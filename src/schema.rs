use serde::Serialize;

pub const TOOL_SCHEMA_VERSION: &str = "1.0.0";
pub const COMMENT_SCHEMA_VERSION: &str = "1.0";
pub const MIN_COMPATIBLE_TOOL_SCHEMA_VERSION: &str = "1.0.0";
pub const SKILL_API_CONTRACT_VERSION: &str = "2026.03";

#[derive(Debug, Clone, Serialize)]
pub struct SchemaInfoResponse {
    pub status: String,
    pub tool_schema_version: String,
    pub min_compatible_version: String,
    pub comment_schema_version: String,
    pub compatibility_guarantees: Vec<String>,
    pub breaking_change_policy: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct SkillApiContractResponse {
    pub status: String,
    pub contract_version: String,
    pub tool_schema_version: String,
    pub min_compatible_tool_schema_version: String,
    pub comment_schema_version: String,
    pub stability_tier: String,
    pub supported_commands: Vec<String>,
    pub supported_mcp_tools: Vec<String>,
    pub compatibility_contract: Vec<String>,
}

pub fn schema_info() -> SchemaInfoResponse {
    SchemaInfoResponse {
        status: "success".to_string(),
        tool_schema_version: TOOL_SCHEMA_VERSION.to_string(),
        min_compatible_version: MIN_COMPATIBLE_TOOL_SCHEMA_VERSION.to_string(),
        comment_schema_version: COMMENT_SCHEMA_VERSION.to_string(),
        compatibility_guarantees: vec![
            "Existing response fields remain stable across minor versions.".to_string(),
            "New fields are additive unless the major version changes.".to_string(),
            "Classic Word comment task payloads remain backward compatible within the 1.x line."
                .to_string(),
        ],
        breaking_change_policy:
            "Breaking changes require a major version bump plus migration guidance.".to_string(),
    }
}

pub fn skill_api_contract() -> SkillApiContractResponse {
    SkillApiContractResponse {
        status: "success".to_string(),
        contract_version: SKILL_API_CONTRACT_VERSION.to_string(),
        tool_schema_version: TOOL_SCHEMA_VERSION.to_string(),
        min_compatible_tool_schema_version: MIN_COMPATIBLE_TOOL_SCHEMA_VERSION.to_string(),
        comment_schema_version: COMMENT_SCHEMA_VERSION.to_string(),
        stability_tier: "stable".to_string(),
        supported_commands: vec![
            "inspect-document".to_string(),
            "extract-text".to_string(),
            "scan-agent-comments".to_string(),
            "resolve-agent-comment-context".to_string(),
            "plan-agent-comment".to_string(),
            "add-agent-comment".to_string(),
            "convert-to-docx".to_string(),
            "doctor".to_string(),
            "init".to_string(),
            "schema-info".to_string(),
            "skill-api-contract".to_string(),
            "mcp-stdio".to_string(),
        ],
        supported_mcp_tools: vec![
            "inspect_document".to_string(),
            "extract_text".to_string(),
            "scan_agent_comments".to_string(),
            "resolve_agent_comment_context".to_string(),
            "plan_agent_comment".to_string(),
            "add_agent_comment".to_string(),
            "convert_to_docx".to_string(),
            "doctor_environment".to_string(),
            "schema_info".to_string(),
            "skill_api_contract".to_string(),
        ],
        compatibility_contract: vec![
            "Task payloads always include task_id, comment_id, selected_text, and instruction."
                .to_string(),
            "Add-agent-comment stays non-destructive by requiring an explicit output_path."
                .to_string(),
            "Plan-agent-comment never mutates the source document and can be used as a mandatory preflight."
                .to_string(),
            "MCP tool names remain stable within the 1.x contract line.".to_string(),
        ],
    }
}
