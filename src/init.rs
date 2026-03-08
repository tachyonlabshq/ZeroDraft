use anyhow::{Context, Result, bail};
use serde::Serialize;
use serde_json::{Value, json};
use std::env;
use std::fs;
use std::path::{Path, PathBuf};

#[derive(Debug, Clone, Serialize)]
pub struct InitReport {
    pub status: String,
    pub project_dir: String,
    pub binary_path: String,
    pub opencode_config_path: String,
    pub skill_stub_path: String,
    pub files_written: Vec<String>,
    pub notes: Vec<String>,
}

pub fn init_project<P: AsRef<Path>>(
    project_dir: P,
    explicit_binary: Option<&Path>,
    force: bool,
) -> Result<InitReport> {
    let project_dir = project_dir.as_ref();
    if !project_dir.exists() {
        bail!(
            "project directory does not exist: {}",
            project_dir.display()
        );
    }
    if !project_dir.is_dir() {
        bail!("project path is not a directory: {}", project_dir.display());
    }

    let binary_path = resolve_binary_path(explicit_binary)?;
    let project_dir = project_dir
        .canonicalize()
        .unwrap_or_else(|_| project_dir.to_path_buf());
    let zerodraft_dir = project_dir.join(".zerodraft");
    fs::create_dir_all(&zerodraft_dir).with_context(|| {
        format!(
            "failed to create ZeroDraft project directory {}",
            zerodraft_dir.display()
        )
    })?;

    let opencode_config_path = project_dir.join("opencode.json");
    let skill_stub_path = zerodraft_dir.join("SKILL.md");
    let mut files_written = Vec::new();
    let mut notes = Vec::new();

    upsert_opencode_config(&opencode_config_path, &binary_path)?;
    files_written.push(opencode_config_path.display().to_string());

    write_skill_stub(&skill_stub_path, &binary_path, force)?;
    files_written.push(skill_stub_path.display().to_string());

    notes.push(
        "project opencode.json now contains or updates an MCP entry named `zerodraft`".to_string(),
    );
    notes.push(
        "use `zerodraft doctor --pretty` after installation to validate runtime prerequisites"
            .to_string(),
    );

    Ok(InitReport {
        status: "success".to_string(),
        project_dir: project_dir.display().to_string(),
        binary_path: binary_path.display().to_string(),
        opencode_config_path: opencode_config_path.display().to_string(),
        skill_stub_path: skill_stub_path.display().to_string(),
        files_written,
        notes,
    })
}

fn resolve_binary_path(explicit_binary: Option<&Path>) -> Result<PathBuf> {
    let candidate = if let Some(path) = explicit_binary {
        path.to_path_buf()
    } else {
        env::current_exe().context("failed to resolve current executable path")?
    };
    if !candidate.exists() {
        bail!("ZeroDraft binary does not exist: {}", candidate.display());
    }
    Ok(candidate.canonicalize().unwrap_or(candidate))
}

fn upsert_opencode_config(path: &Path, binary_path: &Path) -> Result<()> {
    let mut root = if path.exists() {
        let content = fs::read_to_string(path)
            .with_context(|| format!("failed to read {}", path.display()))?;
        serde_json::from_str::<Value>(&content)
            .with_context(|| format!("failed to parse {}", path.display()))?
    } else {
        json!({
            "$schema": "https://opencode.ai/config.json"
        })
    };

    let root_obj = root
        .as_object_mut()
        .ok_or_else(|| anyhow::anyhow!("opencode config root must be a JSON object"))?;
    if !root_obj.contains_key("$schema") {
        root_obj.insert(
            "$schema".to_string(),
            Value::String("https://opencode.ai/config.json".to_string()),
        );
    }

    let mcp_value = root_obj
        .entry("mcp".to_string())
        .or_insert_with(|| Value::Object(Default::default()));
    let mcp_obj = mcp_value
        .as_object_mut()
        .ok_or_else(|| anyhow::anyhow!("opencode config field `mcp` must be a JSON object"))?;

    mcp_obj.insert(
        "zerodraft".to_string(),
        json!({
            "type": "local",
            "command": [
                binary_path.display().to_string(),
                "mcp-stdio"
            ],
            "enabled": true
        }),
    );

    let rendered = serde_json::to_string_pretty(&root)? + "\n";
    fs::write(path, rendered).with_context(|| format!("failed to write {}", path.display()))?;
    Ok(())
}

fn write_skill_stub(path: &Path, binary_path: &Path, force: bool) -> Result<()> {
    if path.exists() && !force {
        return Ok(());
    }

    let content = format!(
        "# ZeroDraft Project Skill\n\n\
This project is configured to use ZeroDraft through MCP.\n\n\
## Binary\n\n\
- `{}`\n\n\
## Recommended workflow\n\n\
1. Use `inspect_document` before operating on an unfamiliar DOCX.\n\
2. Use `scan_agent_comments` to discover Word comments tagged with `@Agent`.\n\
3. Use `resolve_agent_comment_context` before rewriting a highlighted range.\n\
4. Use `plan_agent_comment` before any new writeback so the target range and XML side effects are explicit.\n\
5. Use `add_agent_comment` when you need to create a targeted follow-up instruction.\n\
6. Use `doctor` if `.doc` conversion or MCP setup looks unhealthy.\n",
        binary_path.display()
    );
    fs::write(path, content).with_context(|| format!("failed to write {}", path.display()))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;

    #[test]
    fn init_project_creates_opencode_config() -> Result<()> {
        let tmp = tempdir()?;
        let fake_binary = tmp.path().join("zerodraft");
        fs::write(&fake_binary, "binary")?;

        let report = init_project(tmp.path(), Some(&fake_binary), false)?;
        let config: Value =
            serde_json::from_str(&fs::read_to_string(&report.opencode_config_path)?)?;

        assert_eq!(config["mcp"]["zerodraft"]["enabled"], Value::Bool(true));
        assert!(Path::new(&report.skill_stub_path).exists());
        Ok(())
    }
}
