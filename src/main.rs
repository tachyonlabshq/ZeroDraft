use anyhow::Result;
use clap::Parser;
use zerodraft::{
    AddAgentCommentRequest, Cli, Commands, convert_to_docx, doctor_environment, extract_text,
    init_project, inspect_document, plan_agent_comment, print_json, resolve_agent_comment_context,
    run_mcp_stdio, scan_agent_comments, schema_info, skill_api_contract,
};

fn main() -> Result<()> {
    let cli = Cli::parse();

    match cli.command {
        Commands::InspectDocument {
            document_path,
            pretty,
        } => print_json(&inspect_document(&document_path)?, pretty),
        Commands::ExtractText {
            document_path,
            max_paragraphs,
            pretty,
        } => print_json(&extract_text(&document_path, max_paragraphs)?, pretty),
        Commands::ScanAgentComments {
            document_path,
            pretty,
        } => print_json(&scan_agent_comments(&document_path)?, pretty),
        Commands::ResolveAgentCommentContext {
            document_path,
            task_id,
            window_radius,
            pretty,
        } => print_json(
            &resolve_agent_comment_context(&document_path, &task_id, window_radius)?,
            pretty,
        ),
        Commands::PlanAgentComment {
            document_path,
            comment_text,
            search_text,
            occurrence,
            paragraph_index,
            start_char,
            end_char,
            pretty,
        } => {
            let request = AddAgentCommentRequest {
                document_path,
                output_path: std::path::PathBuf::from("__dry_run__.docx"),
                comment_text,
                author: None,
                search_text,
                occurrence,
                paragraph_index,
                start_char,
                end_char,
            };
            print_json(&plan_agent_comment(request)?, pretty)
        }
        Commands::AddAgentComment {
            document_path,
            output_path,
            comment_text,
            author,
            search_text,
            occurrence,
            paragraph_index,
            start_char,
            end_char,
            pretty,
        } => {
            let request = AddAgentCommentRequest {
                document_path,
                output_path,
                comment_text,
                author,
                search_text,
                occurrence,
                paragraph_index,
                start_char,
                end_char,
            };
            print_json(&zerodraft::add_agent_comment(request)?, pretty)
        }
        Commands::ConvertToDocx {
            input_path,
            output_path,
            pretty,
        } => print_json(&convert_to_docx(&input_path, &output_path)?, pretty),
        Commands::Doctor { pretty } => print_json(&doctor_environment()?, pretty),
        Commands::Init {
            project_dir,
            binary_path,
            force,
            pretty,
        } => print_json(
            &init_project(&project_dir, binary_path.as_deref(), force)?,
            pretty,
        ),
        Commands::SchemaInfo { pretty } => print_json(&schema_info(), pretty),
        Commands::SkillApiContract { pretty } => print_json(&skill_api_contract(), pretty),
        Commands::McpStdio { pretty } => run_mcp_stdio(pretty),
    }
}
