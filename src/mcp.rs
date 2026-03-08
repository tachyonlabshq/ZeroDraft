use anyhow::{Context, Result, anyhow, bail};
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};
use std::io::{self, Write};
use std::path::PathBuf;

use crate::{
    AddAgentCommentRequest, convert_to_docx, doctor_environment, extract_text, inspect_document,
    plan_agent_comment, resolve_agent_comment_context, scan_agent_comments, schema_info,
    skill_api_contract,
};

#[derive(Debug, Clone, Deserialize)]
pub struct McpRequest {
    pub id: Option<Value>,
    pub method: String,
    pub params: Option<Value>,
}

#[derive(Debug, Clone, Serialize)]
pub struct McpResponse {
    pub jsonrpc: String,
    pub id: Option<Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub result: Option<Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error: Option<McpErrorPayload>,
}

#[derive(Debug, Clone, Serialize)]
pub struct McpErrorPayload {
    pub code: i64,
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub data: Option<Value>,
}

#[derive(Debug, Clone, Serialize)]
struct McpToolDescriptor {
    name: String,
    description: String,
    #[serde(rename = "inputSchema")]
    input_schema: Value,
}

#[derive(Debug, Clone, Deserialize)]
struct ToolsCallParams {
    name: String,
    arguments: Option<Value>,
}

pub fn run_mcp_stdio(pretty: bool) -> Result<()> {
    let mut stdout = io::stdout();
    let stdin = io::stdin();
    let stream = serde_json::Deserializer::from_reader(stdin.lock()).into_iter::<Value>();

    for parsed in stream {
        let response = match parsed {
            Ok(raw) => match serde_json::from_value::<McpRequest>(raw) {
                Ok(request) => handle_mcp_request(request),
                Err(err) => McpResponse {
                    jsonrpc: "2.0".to_string(),
                    id: None,
                    result: None,
                    error: Some(McpErrorPayload {
                        code: -32700,
                        message: format!("invalid MCP request JSON: {err}"),
                        data: None,
                    }),
                },
            },
            Err(err) => McpResponse {
                jsonrpc: "2.0".to_string(),
                id: None,
                result: None,
                error: Some(McpErrorPayload {
                    code: -32700,
                    message: format!("invalid MCP request stream: {err}"),
                    data: None,
                }),
            },
        };

        let serialized = if pretty {
            serde_json::to_string_pretty(&response)?
        } else {
            serde_json::to_string(&response)?
        };
        writeln!(stdout, "{serialized}").context("failed to write MCP response line")?;
        stdout.flush().context("failed to flush MCP stdout")?;
    }

    Ok(())
}

pub fn handle_mcp_request(request: McpRequest) -> McpResponse {
    let result = match request.method.as_str() {
        "initialize" => Ok(json!({
            "protocolVersion": "2024-11-05",
            "capabilities": { "tools": {} },
            "serverInfo": {
                "name": "zerodraft",
                "version": "0.1.0"
            }
        })),
        "notifications/initialized" | "notifications/cancelled" => Ok(json!(null)),
        "tools/list" => Ok(json!({ "tools": list_tools() })),
        "tools/call" => call_tool(request.params.clone()),
        unsupported => Err(anyhow!("unsupported method '{unsupported}'")),
    };

    match result {
        Ok(value) => McpResponse {
            jsonrpc: "2.0".to_string(),
            id: request.id,
            result: Some(value),
            error: None,
        },
        Err(err) => McpResponse {
            jsonrpc: "2.0".to_string(),
            id: request.id,
            result: None,
            error: Some(McpErrorPayload {
                code: -32000,
                message: err.to_string(),
                data: None,
            }),
        },
    }
}

fn call_tool(params: Option<Value>) -> Result<Value> {
    let params = params.ok_or_else(|| anyhow!("tools/call requires params"))?;
    let parsed: ToolsCallParams =
        serde_json::from_value(params).context("failed to parse tools/call params")?;
    let arguments = parsed.arguments.unwrap_or_else(|| json!({}));
    let args = arguments
        .as_object()
        .ok_or_else(|| anyhow!("tool arguments must be an object"))?;

    let result = match parsed.name.as_str() {
        "inspect_document" => {
            let path = required_string_from_scopes(args, "document_path", &["path", "file_path"])?;
            serde_json::to_value(inspect_document(path)?)?
        }
        "extract_text" => {
            let path = required_string_from_scopes(args, "document_path", &["path", "file_path"])?;
            let max_paragraphs = args
                .get("max_paragraphs")
                .and_then(Value::as_u64)
                .map(|value| value as usize);
            serde_json::to_value(extract_text(path, max_paragraphs)?)?
        }
        "scan_agent_comments" => {
            let path = required_string_from_scopes(args, "document_path", &["path", "file_path"])?;
            serde_json::to_value(scan_agent_comments(path)?)?
        }
        "resolve_agent_comment_context" => {
            let path = required_string_from_scopes(args, "document_path", &["path", "file_path"])?;
            let task_id = required_string(args, "task_id")?;
            let window_radius = args
                .get("window_radius")
                .and_then(Value::as_u64)
                .unwrap_or(2) as usize;
            serde_json::to_value(resolve_agent_comment_context(path, task_id, window_radius)?)?
        }
        "plan_agent_comment" => {
            let document_path = PathBuf::from(required_string_from_scopes(
                args,
                "document_path",
                &["path"],
            )?);
            let comment_text = required_string(args, "comment_text")?.to_string();
            let search_text = optional_string(args, "search_text").map(str::to_string);
            let occurrence = args.get("occurrence").and_then(Value::as_u64).unwrap_or(1) as usize;
            let paragraph_index = args
                .get("paragraph_index")
                .and_then(Value::as_u64)
                .map(|value| value as usize);
            let start_char = args
                .get("start_char")
                .and_then(Value::as_u64)
                .map(|value| value as usize);
            let end_char = args
                .get("end_char")
                .and_then(Value::as_u64)
                .map(|value| value as usize);

            serde_json::to_value(plan_agent_comment(AddAgentCommentRequest {
                document_path,
                output_path: PathBuf::from("__dry_run__.docx"),
                comment_text,
                author: None,
                search_text,
                occurrence,
                paragraph_index,
                start_char,
                end_char,
            })?)?
        }
        "add_agent_comment" => {
            let document_path = PathBuf::from(required_string_from_scopes(
                args,
                "document_path",
                &["path"],
            )?);
            let output_path = PathBuf::from(required_string(args, "output_path")?);
            let comment_text = required_string(args, "comment_text")?.to_string();
            let author = optional_string(args, "author").map(str::to_string);
            let search_text = optional_string(args, "search_text").map(str::to_string);
            let occurrence = args.get("occurrence").and_then(Value::as_u64).unwrap_or(1) as usize;
            let paragraph_index = args
                .get("paragraph_index")
                .and_then(Value::as_u64)
                .map(|value| value as usize);
            let start_char = args
                .get("start_char")
                .and_then(Value::as_u64)
                .map(|value| value as usize);
            let end_char = args
                .get("end_char")
                .and_then(Value::as_u64)
                .map(|value| value as usize);

            serde_json::to_value(crate::add_agent_comment(AddAgentCommentRequest {
                document_path,
                output_path,
                comment_text,
                author,
                search_text,
                occurrence,
                paragraph_index,
                start_char,
                end_char,
            })?)?
        }
        "convert_to_docx" => {
            let input_path = required_string(args, "input_path")?;
            let output_path = required_string(args, "output_path")?;
            serde_json::to_value(convert_to_docx(input_path, output_path)?)?
        }
        "doctor_environment" => serde_json::to_value(doctor_environment()?)?,
        "schema_info" => serde_json::to_value(schema_info())?,
        "skill_api_contract" => serde_json::to_value(skill_api_contract())?,
        unsupported => bail!("unsupported tool '{unsupported}'"),
    };

    Ok(json!({
        "content": [
            {
                "type": "text",
                "text": serde_json::to_string_pretty(&result)?
            }
        ],
        "structuredContent": result
    }))
}

fn list_tools() -> Vec<McpToolDescriptor> {
    vec![
        McpToolDescriptor {
            name: "inspect_document".to_string(),
            description: "Inspect a DOCX package and return paragraph, table, and comment counts."
                .to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "document_path": { "type": "string" }
                },
                "required": ["document_path"]
            }),
        },
        McpToolDescriptor {
            name: "extract_text".to_string(),
            description: "Extract visible paragraph text from a DOCX document.".to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "document_path": { "type": "string" },
                    "max_paragraphs": { "type": "integer", "minimum": 1 }
                },
                "required": ["document_path"]
            }),
        },
        McpToolDescriptor {
            name: "scan_agent_comments".to_string(),
            description:
                "Scan Word comments for @Agent instructions and return the highlighted target text."
                    .to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "document_path": { "type": "string" }
                },
                "required": ["document_path"]
            }),
        },
        McpToolDescriptor {
            name: "resolve_agent_comment_context".to_string(),
            description:
                "Return surrounding paragraph context for a task produced by scan_agent_comments."
                    .to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "document_path": { "type": "string" },
                    "task_id": { "type": "string" },
                    "window_radius": { "type": "integer", "minimum": 0 }
                },
                "required": ["document_path", "task_id"]
            }),
        },
        McpToolDescriptor {
            name: "plan_agent_comment".to_string(),
            description:
                "Dry-run a targeted Word comment insertion and return the selected text plus expected OOXML side effects."
                    .to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "document_path": { "type": "string" },
                    "comment_text": { "type": "string" },
                    "search_text": { "type": "string" },
                    "occurrence": { "type": "integer", "minimum": 1 },
                    "paragraph_index": { "type": "integer", "minimum": 0 },
                    "start_char": { "type": "integer", "minimum": 0 },
                    "end_char": { "type": "integer", "minimum": 1 }
                },
                "required": ["document_path", "comment_text"]
            }),
        },
        McpToolDescriptor {
            name: "add_agent_comment".to_string(),
            description:
                "Insert a classic Word comment and highlight the targeted text range for agent follow-up."
                    .to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "document_path": { "type": "string" },
                    "output_path": { "type": "string" },
                    "comment_text": { "type": "string" },
                    "author": { "type": "string" },
                    "search_text": { "type": "string" },
                    "occurrence": { "type": "integer", "minimum": 1 },
                    "paragraph_index": { "type": "integer", "minimum": 0 },
                    "start_char": { "type": "integer", "minimum": 0 },
                    "end_char": { "type": "integer", "minimum": 1 }
                },
                "required": ["document_path", "output_path", "comment_text"]
            }),
        },
        McpToolDescriptor {
            name: "convert_to_docx".to_string(),
            description:
                "Convert a legacy .doc file to .docx using LibreOffice headless mode."
                    .to_string(),
            input_schema: json!({
                "type": "object",
                "properties": {
                    "input_path": { "type": "string" },
                    "output_path": { "type": "string" }
                },
                "required": ["input_path", "output_path"]
            }),
        },
        McpToolDescriptor {
            name: "doctor_environment".to_string(),
            description: "Inspect the local runtime environment used by ZeroDraft.".to_string(),
            input_schema: json!({ "type": "object", "properties": {} }),
        },
        McpToolDescriptor {
            name: "schema_info".to_string(),
            description: "Return the stable ZeroDraft schema versions.".to_string(),
            input_schema: json!({ "type": "object", "properties": {} }),
        },
        McpToolDescriptor {
            name: "skill_api_contract".to_string(),
            description: "Return the stable ZeroDraft CLI and MCP contract surface.".to_string(),
            input_schema: json!({ "type": "object", "properties": {} }),
        },
    ]
}

fn required_string<'a>(args: &'a serde_json::Map<String, Value>, key: &str) -> Result<&'a str> {
    args.get(key)
        .and_then(Value::as_str)
        .ok_or_else(|| anyhow!("missing required argument '{key}'"))
}

fn optional_string<'a>(args: &'a serde_json::Map<String, Value>, key: &str) -> Option<&'a str> {
    args.get(key).and_then(Value::as_str)
}

fn required_string_from_scopes<'a>(
    args: &'a serde_json::Map<String, Value>,
    primary: &str,
    aliases: &[&str],
) -> Result<&'a str> {
    if let Some(value) = optional_string(args, primary) {
        return Ok(value);
    }
    for alias in aliases {
        if let Some(value) = optional_string(args, alias) {
            return Ok(value);
        }
    }
    Err(anyhow!("missing required argument '{primary}'"))
}
