use anyhow::{Context, Result, anyhow, bail};
use clap::{Parser, Subcommand};
use serde::Serialize;
use std::collections::{BTreeMap, HashMap};
use std::fs;
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};
use std::process::Command;
use xmltree::{Element, EmitterConfig, XMLNode};
use zip::{CompressionMethod, ZipArchive, ZipWriter, write::SimpleFileOptions};

mod doctor;
mod init;
mod mcp;
mod schema;

pub use doctor::*;
pub use init::*;
pub use mcp::*;
pub use schema::*;

const DOCUMENT_XML_PATH: &str = "word/document.xml";
const COMMENTS_XML_PATH: &str = "word/comments.xml";
const DOCUMENT_RELS_XML_PATH: &str = "word/_rels/document.xml.rels";
const CONTENT_TYPES_XML_PATH: &str = "[Content_Types].xml";
const COMMENTS_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const COMMENTS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

#[derive(Debug, Parser)]
#[command(
    name = "zerodraft",
    version,
    about = "Rust-native Word and DOCX skill for AI agents"
)]
pub struct Cli {
    #[command(subcommand)]
    pub command: Commands,
}

#[derive(Debug, Subcommand)]
pub enum Commands {
    InspectDocument {
        document_path: PathBuf,
        #[arg(long)]
        pretty: bool,
    },
    ExtractText {
        document_path: PathBuf,
        #[arg(long)]
        max_paragraphs: Option<usize>,
        #[arg(long)]
        pretty: bool,
    },
    ScanAgentComments {
        document_path: PathBuf,
        #[arg(long)]
        pretty: bool,
    },
    ResolveAgentCommentContext {
        document_path: PathBuf,
        task_id: String,
        #[arg(long, default_value_t = 2)]
        window_radius: usize,
        #[arg(long)]
        pretty: bool,
    },
    AddAgentComment {
        document_path: PathBuf,
        output_path: PathBuf,
        #[arg(long)]
        comment_text: String,
        #[arg(long)]
        author: Option<String>,
        #[arg(long)]
        search_text: Option<String>,
        #[arg(long, default_value_t = 1)]
        occurrence: usize,
        #[arg(long)]
        paragraph_index: Option<usize>,
        #[arg(long)]
        start_char: Option<usize>,
        #[arg(long)]
        end_char: Option<usize>,
        #[arg(long)]
        pretty: bool,
    },
    ConvertToDocx {
        input_path: PathBuf,
        output_path: PathBuf,
        #[arg(long)]
        pretty: bool,
    },
    Doctor {
        #[arg(long)]
        pretty: bool,
    },
    Init {
        project_dir: PathBuf,
        #[arg(long)]
        binary_path: Option<PathBuf>,
        #[arg(long)]
        force: bool,
        #[arg(long)]
        pretty: bool,
    },
    SchemaInfo {
        #[arg(long)]
        pretty: bool,
    },
    SkillApiContract {
        #[arg(long)]
        pretty: bool,
    },
    McpStdio {
        #[arg(long)]
        pretty: bool,
    },
}

#[derive(Debug, Clone, Serialize)]
pub struct DocumentInspection {
    pub status: String,
    pub document_path: String,
    pub paragraph_count: usize,
    pub non_empty_paragraph_count: usize,
    pub table_count: usize,
    pub comment_count: usize,
    pub agent_comment_count: usize,
    pub has_comments_part: bool,
}

#[derive(Debug, Clone, Serialize)]
pub struct ParagraphText {
    pub index: usize,
    pub text: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct ExtractTextResponse {
    pub status: String,
    pub document_path: String,
    pub paragraph_count: usize,
    pub paragraphs: Vec<ParagraphText>,
}

#[derive(Debug, Clone, Serialize)]
pub struct AgentCommentScanReport {
    pub status: String,
    pub document_path: String,
    pub comment_count: usize,
    pub task_count: usize,
    pub tasks: Vec<AgentCommentTask>,
}

#[derive(Debug, Clone, Serialize)]
pub struct AgentCommentTask {
    pub task_id: String,
    pub comment_id: String,
    pub author: Option<String>,
    pub paragraph_start: usize,
    pub paragraph_end: usize,
    pub selected_text: String,
    pub raw_comment_text: String,
    pub instruction: String,
    pub trigger: String,
}

#[derive(Debug, Clone, Serialize)]
pub struct AgentCommentContext {
    pub status: String,
    pub document_path: String,
    pub window_radius: usize,
    pub task: AgentCommentTask,
    pub context_paragraphs: Vec<ParagraphText>,
}

#[derive(Debug, Clone)]
pub struct AddAgentCommentRequest {
    pub document_path: PathBuf,
    pub output_path: PathBuf,
    pub comment_text: String,
    pub author: Option<String>,
    pub search_text: Option<String>,
    pub occurrence: usize,
    pub paragraph_index: Option<usize>,
    pub start_char: Option<usize>,
    pub end_char: Option<usize>,
}

#[derive(Debug, Clone, Serialize)]
pub struct AddAgentCommentReport {
    pub status: String,
    pub document_path: String,
    pub output_path: String,
    pub comment_id: String,
    pub paragraph_index: usize,
    pub start_char: usize,
    pub end_char: usize,
    pub selected_text: String,
    pub comments_part_created: bool,
    pub relationship_created: bool,
}

#[derive(Debug, Clone, Serialize)]
pub struct ConvertToDocxReport {
    pub status: String,
    pub input_path: String,
    pub output_path: String,
    pub converter: String,
    pub notes: Vec<String>,
}

#[derive(Debug, Clone)]
struct CommentRecord {
    author: Option<String>,
    text: String,
}

#[derive(Debug, Clone)]
struct CommentRangeCapture {
    id: String,
    paragraph_start: usize,
    paragraph_end: usize,
    selected_text: String,
}

#[derive(Debug, Clone)]
struct TraversalState {
    paragraphs: Vec<String>,
    current_paragraph: String,
    current_paragraph_index: usize,
    table_count: usize,
    active_comments: Vec<String>,
    captures: BTreeMap<String, CommentRangeCapture>,
}

#[derive(Debug, Clone)]
struct ParagraphSelection {
    paragraph_index: usize,
    start_char: usize,
    end_char: usize,
    selected_text: String,
}

#[derive(Debug, Clone)]
struct DocxPackage {
    entries: BTreeMap<String, Vec<u8>>,
}

impl DocxPackage {
    fn open(path: &Path) -> Result<Self> {
        let bytes = fs::read(path)
            .with_context(|| format!("failed to read document {}", path.display()))?;
        let reader = Cursor::new(bytes);
        let mut zip = ZipArchive::new(reader).context("failed to open DOCX zip archive")?;
        let mut entries = BTreeMap::new();
        for i in 0..zip.len() {
            let mut file = zip.by_index(i)?;
            let name = file.name().to_string();
            let mut data = Vec::new();
            file.read_to_end(&mut data)?;
            entries.insert(name, data);
        }
        if !entries.contains_key(DOCUMENT_XML_PATH) {
            bail!("unsupported document: missing {DOCUMENT_XML_PATH}");
        }
        if !entries.contains_key(CONTENT_TYPES_XML_PATH) {
            bail!("unsupported document: missing {CONTENT_TYPES_XML_PATH}");
        }
        Ok(Self { entries })
    }

    fn read_xml(&self, path: &str) -> Result<Option<Element>> {
        let Some(bytes) = self.entries.get(path) else {
            return Ok(None);
        };
        let cursor = Cursor::new(bytes);
        Ok(Some(Element::parse(cursor).with_context(|| {
            format!("failed to parse XML part {path}")
        })?))
    }

    fn write_xml(&mut self, path: &str, element: &Element) -> Result<()> {
        let mut buf = Vec::new();
        element
            .write_with_config(
                &mut buf,
                EmitterConfig::new()
                    .perform_indent(false)
                    .write_document_declaration(true),
            )
            .with_context(|| format!("failed to serialize XML part {path}"))?;
        self.entries.insert(path.to_string(), buf);
        Ok(())
    }

    fn contains(&self, path: &str) -> bool {
        self.entries.contains_key(path)
    }

    fn write_to(&self, path: &Path) -> Result<()> {
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent)
                .with_context(|| format!("failed to create {}", parent.display()))?;
        }
        let file = fs::File::create(path)
            .with_context(|| format!("failed to create {}", path.display()))?;
        let mut writer = ZipWriter::new(file);
        let options = SimpleFileOptions::default()
            .compression_method(CompressionMethod::Deflated)
            .unix_permissions(0o644);
        for (name, data) in &self.entries {
            writer.start_file(name, options)?;
            writer.write_all(data)?;
        }
        writer.finish()?;
        Ok(())
    }
}

pub fn inspect_document<P: AsRef<Path>>(path: P) -> Result<DocumentInspection> {
    let path = path.as_ref();
    let package = DocxPackage::open(path)?;
    let document = package
        .read_xml(DOCUMENT_XML_PATH)?
        .ok_or_else(|| anyhow!("missing {DOCUMENT_XML_PATH}"))?;
    let comments = read_comments_map(&package)?;
    let traversed = traverse_document(&document);
    let agent_comment_count = comments
        .values()
        .filter(|record| contains_agent_trigger(&record.text))
        .count();

    Ok(DocumentInspection {
        status: "success".to_string(),
        document_path: path.display().to_string(),
        paragraph_count: traversed.paragraphs.len(),
        non_empty_paragraph_count: traversed
            .paragraphs
            .iter()
            .filter(|text| !text.trim().is_empty())
            .count(),
        table_count: traversed.table_count,
        comment_count: comments.len(),
        agent_comment_count,
        has_comments_part: package.contains(COMMENTS_XML_PATH),
    })
}

pub fn extract_text<P: AsRef<Path>>(
    path: P,
    max_paragraphs: Option<usize>,
) -> Result<ExtractTextResponse> {
    let path = path.as_ref();
    let package = DocxPackage::open(path)?;
    let document = package
        .read_xml(DOCUMENT_XML_PATH)?
        .ok_or_else(|| anyhow!("missing {DOCUMENT_XML_PATH}"))?;
    let traversed = traverse_document(&document);
    let paragraphs = traversed
        .paragraphs
        .iter()
        .enumerate()
        .take(max_paragraphs.unwrap_or(usize::MAX))
        .map(|(index, text)| ParagraphText {
            index,
            text: text.clone(),
        })
        .collect::<Vec<_>>();

    Ok(ExtractTextResponse {
        status: "success".to_string(),
        document_path: path.display().to_string(),
        paragraph_count: traversed.paragraphs.len(),
        paragraphs,
    })
}

pub fn scan_agent_comments<P: AsRef<Path>>(path: P) -> Result<AgentCommentScanReport> {
    let path = path.as_ref();
    let package = DocxPackage::open(path)?;
    let document = package
        .read_xml(DOCUMENT_XML_PATH)?
        .ok_or_else(|| anyhow!("missing {DOCUMENT_XML_PATH}"))?;
    let traversed = traverse_document(&document);
    let comments = read_comments_map(&package)?;
    let tasks = traversed
        .captures
        .values()
        .filter_map(|capture| {
            let record = comments.get(&capture.id)?;
            if !contains_agent_trigger(&record.text) {
                return None;
            }
            Some(AgentCommentTask {
                task_id: format!("comment-{}", capture.id),
                comment_id: capture.id.clone(),
                author: record.author.clone(),
                paragraph_start: capture.paragraph_start,
                paragraph_end: capture.paragraph_end,
                selected_text: capture.selected_text.clone(),
                raw_comment_text: record.text.clone(),
                instruction: normalize_agent_instruction(&record.text),
                trigger: "@agent".to_string(),
            })
        })
        .collect::<Vec<_>>();

    Ok(AgentCommentScanReport {
        status: "success".to_string(),
        document_path: path.display().to_string(),
        comment_count: comments.len(),
        task_count: tasks.len(),
        tasks,
    })
}

pub fn resolve_agent_comment_context<P: AsRef<Path>>(
    path: P,
    task_id: &str,
    window_radius: usize,
) -> Result<AgentCommentContext> {
    let path = path.as_ref();
    let scan = scan_agent_comments(path)?;
    let task = scan
        .tasks
        .iter()
        .find(|task| task.task_id == task_id)
        .cloned()
        .ok_or_else(|| anyhow!("task_id not found: {task_id}"))?;
    let extracted = extract_text(path, None)?;
    let start = task.paragraph_start.saturating_sub(window_radius);
    let end = usize::min(
        extracted.paragraph_count.saturating_sub(1),
        task.paragraph_end + window_radius,
    );
    let context_paragraphs = extracted
        .paragraphs
        .into_iter()
        .filter(|paragraph| paragraph.index >= start && paragraph.index <= end)
        .collect::<Vec<_>>();

    Ok(AgentCommentContext {
        status: "success".to_string(),
        document_path: path.display().to_string(),
        window_radius,
        task,
        context_paragraphs,
    })
}

pub fn add_agent_comment(request: AddAgentCommentRequest) -> Result<AddAgentCommentReport> {
    validate_add_comment_request(&request)?;

    let mut package = DocxPackage::open(&request.document_path)?;
    let comments_before = package.contains(COMMENTS_XML_PATH);
    let document = package
        .read_xml(DOCUMENT_XML_PATH)?
        .ok_or_else(|| anyhow!("missing {DOCUMENT_XML_PATH}"))?;
    let extracted = extract_text(&request.document_path, None)?;
    let selection = resolve_selection(&extracted.paragraphs, &request)?;
    let comment_id = next_comment_id(&package)?;

    let mut document = document;
    insert_comment_range(&mut document, &selection, &comment_id)?;
    package.write_xml(DOCUMENT_XML_PATH, &document)?;

    let relationship_created = ensure_comments_relationship(&mut package)?;
    ensure_comments_content_type(&mut package)?;
    append_comment_record(
        &mut package,
        &comment_id,
        request.author.as_deref().unwrap_or("ZeroDraft"),
        &request.comment_text,
    )?;
    package.write_to(&request.output_path)?;

    Ok(AddAgentCommentReport {
        status: "success".to_string(),
        document_path: request.document_path.display().to_string(),
        output_path: request.output_path.display().to_string(),
        comment_id,
        paragraph_index: selection.paragraph_index,
        start_char: selection.start_char,
        end_char: selection.end_char,
        selected_text: selection.selected_text,
        comments_part_created: !comments_before,
        relationship_created,
    })
}

pub fn convert_to_docx<P: AsRef<Path>, Q: AsRef<Path>>(
    input_path: P,
    output_path: Q,
) -> Result<ConvertToDocxReport> {
    let input_path = input_path.as_ref();
    let output_path = output_path.as_ref();
    let output_dir = output_path
        .parent()
        .ok_or_else(|| anyhow!("output_path must have a parent directory"))?;
    fs::create_dir_all(output_dir)
        .with_context(|| format!("failed to create {}", output_dir.display()))?;

    let soffice = find_executable("soffice")
        .or_else(|| find_executable("libreoffice"))
        .ok_or_else(|| {
            anyhow!("LibreOffice was not found on PATH; cannot convert .doc to .docx")
        })?;

    let status = Command::new(&soffice)
        .arg("--headless")
        .arg("--convert-to")
        .arg("docx")
        .arg("--outdir")
        .arg(output_dir)
        .arg(input_path)
        .status()
        .with_context(|| format!("failed to execute {}", soffice.display()))?;
    if !status.success() {
        bail!("conversion command failed with status {status}");
    }

    let converted_name = input_path
        .file_stem()
        .ok_or_else(|| anyhow!("input file must have a file name"))?;
    let produced = output_dir.join(format!("{}.docx", converted_name.to_string_lossy()));
    if !produced.exists() {
        bail!(
            "conversion reported success but no DOCX was produced at {}",
            produced.display()
        );
    }
    if produced != output_path {
        fs::copy(&produced, output_path).with_context(|| {
            format!(
                "failed to copy converted file from {} to {}",
                produced.display(),
                output_path.display()
            )
        })?;
    }

    Ok(ConvertToDocxReport {
        status: "success".to_string(),
        input_path: input_path.display().to_string(),
        output_path: output_path.display().to_string(),
        converter: soffice.display().to_string(),
        notes: vec![
            "Conversion uses LibreOffice headless mode.".to_string(),
            "This path is intended for .doc compatibility when native DOCX operations are preferred."
                .to_string(),
        ],
    })
}

pub fn print_json<T: Serialize>(value: &T, pretty: bool) -> Result<()> {
    let rendered = if pretty {
        serde_json::to_string_pretty(value)?
    } else {
        serde_json::to_string(value)?
    };
    println!("{rendered}");
    Ok(())
}

fn validate_add_comment_request(request: &AddAgentCommentRequest) -> Result<()> {
    if request.comment_text.trim().is_empty() {
        bail!("comment_text must not be empty");
    }

    let has_search = request
        .search_text
        .as_ref()
        .map(|text| !text.is_empty())
        .unwrap_or(false);
    let has_range = request.paragraph_index.is_some()
        || request.start_char.is_some()
        || request.end_char.is_some();

    if has_search == has_range {
        bail!(
            "provide either search_text plus occurrence, or paragraph_index with start_char and end_char"
        );
    }

    if has_range
        && (request.paragraph_index.is_none()
            || request.start_char.is_none()
            || request.end_char.is_none())
    {
        bail!("paragraph_index, start_char, and end_char are all required for explicit ranges");
    }

    Ok(())
}

fn resolve_selection(
    paragraphs: &[ParagraphText],
    request: &AddAgentCommentRequest,
) -> Result<ParagraphSelection> {
    if let Some(search_text) = &request.search_text {
        let mut seen = 0usize;
        for paragraph in paragraphs {
            let mut search_from = 0usize;
            while let Some(idx) = paragraph.text[search_from..].find(search_text) {
                seen += 1;
                let start = char_count(&paragraph.text[..search_from + idx]);
                let end = start + char_count(search_text);
                if seen == request.occurrence {
                    return Ok(ParagraphSelection {
                        paragraph_index: paragraph.index,
                        start_char: start,
                        end_char: end,
                        selected_text: search_text.clone(),
                    });
                }
                search_from += idx + search_text.len();
            }
        }
        bail!(
            "search_text occurrence {} was not found in the document",
            request.occurrence
        );
    }

    let paragraph_index = request
        .paragraph_index
        .ok_or_else(|| anyhow!("missing paragraph_index"))?;
    let start_char = request
        .start_char
        .ok_or_else(|| anyhow!("missing start_char"))?;
    let end_char = request
        .end_char
        .ok_or_else(|| anyhow!("missing end_char"))?;
    let paragraph = paragraphs
        .iter()
        .find(|paragraph| paragraph.index == paragraph_index)
        .ok_or_else(|| anyhow!("paragraph_index out of bounds: {paragraph_index}"))?;
    if start_char >= end_char {
        bail!("start_char must be less than end_char");
    }
    if end_char > char_count(&paragraph.text) {
        bail!(
            "end_char {} exceeds paragraph character length {}",
            end_char,
            char_count(&paragraph.text)
        );
    }
    Ok(ParagraphSelection {
        paragraph_index,
        start_char,
        end_char,
        selected_text: slice_chars(&paragraph.text, start_char, end_char),
    })
}

fn next_comment_id(package: &DocxPackage) -> Result<String> {
    let comments = read_comments_map(package)?;
    let next = comments
        .keys()
        .filter_map(|id| id.parse::<u32>().ok())
        .max()
        .map(|id| id + 1)
        .unwrap_or(0);
    Ok(next.to_string())
}

fn read_comments_map(package: &DocxPackage) -> Result<HashMap<String, CommentRecord>> {
    let Some(comments) = package.read_xml(COMMENTS_XML_PATH)? else {
        return Ok(HashMap::new());
    };
    let mut map = HashMap::new();
    for child in &comments.children {
        let XMLNode::Element(comment) = child else {
            continue;
        };
        if !element_is(comment, "comment") {
            continue;
        }
        let Some(id) = attr_value(comment, &["w:id", "id"]).cloned() else {
            continue;
        };
        let author = attr_value(comment, &["w:author", "author"]).cloned();
        let text = collect_visible_text(comment);
        map.insert(id.clone(), CommentRecord { author, text });
    }
    Ok(map)
}

fn traverse_document(document: &Element) -> TraversalState {
    let mut state = TraversalState {
        paragraphs: Vec::new(),
        current_paragraph: String::new(),
        current_paragraph_index: 0,
        table_count: 0,
        active_comments: Vec::new(),
        captures: BTreeMap::new(),
    };
    traverse_nodes(&document.children, &mut state);
    state
}

fn traverse_nodes(nodes: &[XMLNode], state: &mut TraversalState) {
    for node in nodes {
        let XMLNode::Element(element) = node else {
            continue;
        };
        match element.name.as_str() {
            name if name_is(name, "tbl") => {
                state.table_count += 1;
                traverse_nodes(&element.children, state);
            }
            name if name_is(name, "p") => {
                state.current_paragraph.clear();
                traverse_paragraph(element, state);
                state.paragraphs.push(state.current_paragraph.clone());
                if !state.active_comments.is_empty() {
                    for active in &state.active_comments {
                        if let Some(capture) = state.captures.get_mut(active) {
                            capture.selected_text.push('\n');
                        }
                    }
                }
                state.current_paragraph_index += 1;
            }
            _ => traverse_nodes(&element.children, state),
        }
    }
}

fn traverse_paragraph(paragraph: &Element, state: &mut TraversalState) {
    traverse_textual_nodes(paragraph, state);
}

fn traverse_textual_nodes(element: &Element, state: &mut TraversalState) {
    match element.name.as_str() {
        name if name_is(name, "commentRangeStart") => {
            if let Some(id) = attr_value(element, &["w:id", "id"])
                && !state.active_comments.iter().any(|active| active == id)
            {
                state.active_comments.push(id.clone());
                state.captures.insert(
                    id.clone(),
                    CommentRangeCapture {
                        id: id.clone(),
                        paragraph_start: state.current_paragraph_index,
                        paragraph_end: state.current_paragraph_index,
                        selected_text: String::new(),
                    },
                );
            }
        }
        name if name_is(name, "commentRangeEnd") => {
            if let Some(id) = attr_value(element, &["w:id", "id"]) {
                state.active_comments.retain(|active| active != id);
            }
        }
        name if name_is(name, "t") => {
            let text = text_content(element);
            state.current_paragraph.push_str(&text);
            for active in &state.active_comments {
                if let Some(capture) = state.captures.get_mut(active) {
                    capture.paragraph_end = state.current_paragraph_index;
                    capture.selected_text.push_str(&text);
                }
            }
        }
        name if name_is(name, "tab") => {
            state.current_paragraph.push('\t');
            for active in &state.active_comments {
                if let Some(capture) = state.captures.get_mut(active) {
                    capture.paragraph_end = state.current_paragraph_index;
                    capture.selected_text.push('\t');
                }
            }
        }
        name if name_is(name, "br") || name_is(name, "cr") => {
            state.current_paragraph.push('\n');
            for active in &state.active_comments {
                if let Some(capture) = state.captures.get_mut(active) {
                    capture.paragraph_end = state.current_paragraph_index;
                    capture.selected_text.push('\n');
                }
            }
        }
        _ => {
            for child in &element.children {
                if let XMLNode::Element(child_element) = child {
                    traverse_textual_nodes(child_element, state);
                }
            }
        }
    }
}

fn contains_agent_trigger(text: &str) -> bool {
    text.to_ascii_lowercase().contains("@agent")
}

fn normalize_agent_instruction(text: &str) -> String {
    let lowered = text.to_ascii_lowercase();
    if let Some(index) = lowered.find("@agent") {
        text[index + "@agent".len()..].trim().to_string()
    } else {
        text.trim().to_string()
    }
}

fn insert_comment_range(
    document: &mut Element,
    selection: &ParagraphSelection,
    comment_id: &str,
) -> Result<()> {
    let mut paragraph_counter = 0usize;
    let paragraph =
        find_nth_paragraph_mut(document, selection.paragraph_index, &mut paragraph_counter)
            .ok_or_else(|| anyhow!("paragraph_index out of bounds"))?;
    mutate_paragraph_with_comment(paragraph, selection, comment_id)
}

fn find_nth_paragraph_mut<'a>(
    element: &'a mut Element,
    target: usize,
    counter: &mut usize,
) -> Option<&'a mut Element> {
    if element_is(element, "p") {
        if *counter == target {
            return Some(element);
        }
        *counter += 1;
    }
    for child in &mut element.children {
        if let XMLNode::Element(child_element) = child
            && let Some(found) = find_nth_paragraph_mut(child_element, target, counter)
        {
            return Some(found);
        }
    }
    None
}

fn mutate_paragraph_with_comment(
    paragraph: &mut Element,
    selection: &ParagraphSelection,
    comment_id: &str,
) -> Result<()> {
    let original_children = std::mem::take(&mut paragraph.children);
    let mut rebuilt = Vec::new();
    let mut char_cursor = 0usize;
    let mut inserted_start = false;
    let mut inserted_end = false;

    for child in original_children {
        match child {
            XMLNode::Element(run) if element_is(&run, "r") => {
                if let Some(run_text) = extract_simple_run_text(&run)? {
                    let run_len = char_count(&run_text);
                    let run_start = char_cursor;
                    let run_end = run_start + run_len;
                    let overlap_start = usize::max(selection.start_char, run_start);
                    let overlap_end = usize::min(selection.end_char, run_end);

                    if overlap_start >= overlap_end {
                        rebuilt.push(XMLNode::Element(run));
                    } else {
                        let before = slice_chars(&run_text, 0, overlap_start - run_start);
                        let selected = slice_chars(
                            &run_text,
                            overlap_start - run_start,
                            overlap_end - run_start,
                        );
                        let after =
                            slice_chars(&run_text, overlap_end - run_start, char_count(&run_text));

                        if !before.is_empty() {
                            rebuilt.push(XMLNode::Element(clone_run_with_text(&run, &before)?));
                        }
                        if !inserted_start {
                            rebuilt.push(XMLNode::Element(make_comment_marker(
                                "w:commentRangeStart",
                                comment_id,
                            )));
                            inserted_start = true;
                        }
                        if !selected.is_empty() {
                            rebuilt.push(XMLNode::Element(clone_run_with_text(&run, &selected)?));
                        }
                        if overlap_end == selection.end_char && !inserted_end {
                            rebuilt.push(XMLNode::Element(make_comment_marker(
                                "w:commentRangeEnd",
                                comment_id,
                            )));
                            rebuilt.push(XMLNode::Element(make_comment_reference_run(comment_id)));
                            inserted_end = true;
                        }
                        if !after.is_empty() {
                            rebuilt.push(XMLNode::Element(clone_run_with_text(&run, &after)?));
                        }
                    }

                    char_cursor = run_end;
                } else {
                    rebuilt.push(XMLNode::Element(run));
                }
            }
            other => rebuilt.push(other),
        }
    }

    if !inserted_start || !inserted_end {
        bail!(
            "selection could not be mapped into a commentable run sequence; choose a simpler range or exact search_text"
        );
    }

    paragraph.children = rebuilt;
    Ok(())
}

fn extract_simple_run_text(run: &Element) -> Result<Option<String>> {
    let mut text_nodes = Vec::new();
    collect_text_elements(run, &mut text_nodes);
    if text_nodes.is_empty() {
        return Ok(None);
    }
    if text_nodes.len() > 1 {
        bail!("runs with multiple text nodes are not yet supported for comment insertion");
    }
    Ok(Some(text_content(text_nodes[0])))
}

fn collect_text_elements<'a>(element: &'a Element, out: &mut Vec<&'a Element>) {
    if element_is(element, "t") {
        out.push(element);
        return;
    }
    for child in &element.children {
        if let XMLNode::Element(child_element) = child {
            collect_text_elements(child_element, out);
        }
    }
}

fn clone_run_with_text(run: &Element, replacement_text: &str) -> Result<Element> {
    let mut cloned = run.clone();
    let mut replaced = false;
    replace_first_text_node(&mut cloned, replacement_text, &mut replaced);
    if !replaced {
        bail!("run did not contain a replaceable w:t node");
    }
    Ok(cloned)
}

fn replace_first_text_node(element: &mut Element, replacement_text: &str, replaced: &mut bool) {
    if *replaced {
        return;
    }
    if element_is(element, "t") {
        element.children.clear();
        element
            .children
            .push(XMLNode::Text(replacement_text.to_string()));
        if replacement_text.starts_with(' ') || replacement_text.ends_with(' ') {
            element
                .attributes
                .insert("xml:space".to_string(), "preserve".to_string());
        } else {
            element.attributes.remove("xml:space");
        }
        *replaced = true;
        return;
    }
    for child in &mut element.children {
        if let XMLNode::Element(child_element) = child {
            replace_first_text_node(child_element, replacement_text, replaced);
            if *replaced {
                return;
            }
        }
    }
}

fn make_comment_marker(name: &str, comment_id: &str) -> Element {
    let mut element = Element::new(name);
    element
        .attributes
        .insert("w:id".to_string(), comment_id.to_string());
    element
}

fn make_comment_reference_run(comment_id: &str) -> Element {
    let mut run = Element::new("w:r");
    let mut reference = Element::new("w:commentReference");
    reference
        .attributes
        .insert("w:id".to_string(), comment_id.to_string());
    run.children.push(XMLNode::Element(reference));
    run
}

fn ensure_comments_relationship(package: &mut DocxPackage) -> Result<bool> {
    let mut rels = package
        .read_xml(DOCUMENT_RELS_XML_PATH)?
        .unwrap_or_else(make_document_relationships_root);
    for child in &rels.children {
        let XMLNode::Element(element) = child else {
            continue;
        };
        if element_is(element, "Relationship")
            && attr_value(element, &["Type", "type"])
                .map(|value| value == COMMENTS_REL_TYPE)
                .unwrap_or(false)
        {
            package.write_xml(DOCUMENT_RELS_XML_PATH, &rels)?;
            return Ok(false);
        }
    }

    let next_id = next_relationship_id(&rels);
    let mut relationship = Element::new("Relationship");
    relationship.attributes.insert("Id".to_string(), next_id);
    relationship
        .attributes
        .insert("Type".to_string(), COMMENTS_REL_TYPE.to_string());
    relationship
        .attributes
        .insert("Target".to_string(), "comments.xml".to_string());
    rels.children.push(XMLNode::Element(relationship));
    package.write_xml(DOCUMENT_RELS_XML_PATH, &rels)?;
    Ok(true)
}

fn next_relationship_id(rels: &Element) -> String {
    let next = rels
        .children
        .iter()
        .filter_map(|node| match node {
            XMLNode::Element(element) if element_is(element, "Relationship") => {
                attr_value(element, &["Id", "id"])
                    .and_then(|id| id.strip_prefix("rId"))
                    .and_then(|suffix| suffix.parse::<u32>().ok())
            }
            _ => None,
        })
        .max()
        .map(|id| id + 1)
        .unwrap_or(1);
    format!("rId{next}")
}

fn ensure_comments_content_type(package: &mut DocxPackage) -> Result<()> {
    let mut content_types = package
        .read_xml(CONTENT_TYPES_XML_PATH)?
        .ok_or_else(|| anyhow!("missing {CONTENT_TYPES_XML_PATH}"))?;
    let exists = content_types.children.iter().any(|node| match node {
        XMLNode::Element(element) if element_is(element, "Override") => {
            attr_value(element, &["PartName", "partname"])
                .map(|value| value == "/word/comments.xml")
                .unwrap_or(false)
        }
        _ => false,
    });
    if !exists {
        let mut override_part = Element::new("Override");
        override_part
            .attributes
            .insert("PartName".to_string(), "/word/comments.xml".to_string());
        override_part
            .attributes
            .insert("ContentType".to_string(), COMMENTS_CONTENT_TYPE.to_string());
        content_types.children.push(XMLNode::Element(override_part));
        package.write_xml(CONTENT_TYPES_XML_PATH, &content_types)?;
    }
    Ok(())
}

fn append_comment_record(
    package: &mut DocxPackage,
    comment_id: &str,
    author: &str,
    comment_text: &str,
) -> Result<()> {
    let mut comments = package
        .read_xml(COMMENTS_XML_PATH)?
        .unwrap_or_else(make_comments_root);
    let mut comment = Element::new("w:comment");
    comment
        .attributes
        .insert("w:id".to_string(), comment_id.to_string());
    comment
        .attributes
        .insert("w:author".to_string(), author.to_string());
    comment
        .attributes
        .insert("w:initials".to_string(), initials(author));

    let mut paragraph = Element::new("w:p");
    let mut run = Element::new("w:r");
    let mut text = Element::new("w:t");
    if comment_text.starts_with(' ') || comment_text.ends_with(' ') {
        text.attributes
            .insert("xml:space".to_string(), "preserve".to_string());
    }
    text.children.push(XMLNode::Text(comment_text.to_string()));
    run.children.push(XMLNode::Element(text));
    paragraph.children.push(XMLNode::Element(run));
    comment.children.push(XMLNode::Element(paragraph));
    comments.children.push(XMLNode::Element(comment));
    package.write_xml(COMMENTS_XML_PATH, &comments)?;
    Ok(())
}

fn initials(author: &str) -> String {
    let mut out = String::new();
    for part in author.split_whitespace().take(2) {
        if let Some(ch) = part.chars().next() {
            out.push(ch.to_ascii_uppercase());
        }
    }
    if out.is_empty() {
        "ZD".to_string()
    } else {
        out
    }
}

fn make_comments_root() -> Element {
    let mut root = Element::new("w:comments");
    root.attributes.insert(
        "xmlns:w".to_string(),
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main".to_string(),
    );
    root
}

fn make_document_relationships_root() -> Element {
    let mut root = Element::new("Relationships");
    root.attributes.insert(
        "xmlns".to_string(),
        "http://schemas.openxmlformats.org/package/2006/relationships".to_string(),
    );
    root
}

fn text_content(element: &Element) -> String {
    let mut out = String::new();
    for child in &element.children {
        if let XMLNode::Text(text) = child {
            out.push_str(text);
        }
    }
    out
}

fn collect_visible_text(element: &Element) -> String {
    let mut out = String::new();
    collect_visible_text_into(element, &mut out);
    out
}

fn collect_visible_text_into(element: &Element, out: &mut String) {
    match element.name.as_str() {
        name if name_is(name, "t") => out.push_str(&text_content(element)),
        name if name_is(name, "tab") => out.push('\t'),
        name if name_is(name, "br") || name_is(name, "cr") => out.push('\n'),
        _ => {
            for child in &element.children {
                if let XMLNode::Element(child_element) = child {
                    collect_visible_text_into(child_element, out);
                }
            }
        }
    }
}

fn slice_chars(text: &str, start: usize, end: usize) -> String {
    text.chars()
        .skip(start)
        .take(end - start)
        .collect::<String>()
}

fn char_count(text: &str) -> usize {
    text.chars().count()
}

fn find_executable(name: &str) -> Option<PathBuf> {
    let output = Command::new("which").arg(name).output().ok()?;
    if !output.status.success() {
        return None;
    }
    let path = String::from_utf8_lossy(&output.stdout).trim().to_string();
    if path.is_empty() {
        None
    } else {
        Some(PathBuf::from(path))
    }
}

fn local_name(name: &str) -> &str {
    name.rsplit(':').next().unwrap_or(name)
}

fn name_is(name: &str, expected_local: &str) -> bool {
    local_name(name) == expected_local
}

fn element_is(element: &Element, expected_local: &str) -> bool {
    name_is(&element.name, expected_local)
}

fn attr_value<'a>(element: &'a Element, candidates: &[&str]) -> Option<&'a String> {
    for candidate in candidates {
        if let Some(value) = element.attributes.get(*candidate) {
            return Some(value);
        }
    }
    let candidate_locals = candidates
        .iter()
        .map(|candidate| local_name(candidate))
        .collect::<Vec<_>>();
    element.attributes.iter().find_map(|(key, value)| {
        if candidate_locals
            .iter()
            .any(|candidate| local_name(key) == *candidate)
        {
            Some(value)
        } else {
            None
        }
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;

    #[test]
    fn scan_agent_comments_extracts_highlighted_text() -> Result<()> {
        let tmp = tempdir()?;
        let doc_path = tmp.path().join("sample.docx");
        write_test_docx(
            &doc_path,
            TEST_DOCUMENT_XML_WITH_COMMENT,
            Some(TEST_COMMENTS_XML),
        )?;

        let scan = scan_agent_comments(&doc_path)?;
        assert_eq!(scan.task_count, 1);
        assert_eq!(scan.tasks[0].selected_text, "targeted text");
        assert_eq!(scan.tasks[0].instruction, "rewrite this section");
        Ok(())
    }

    #[test]
    fn add_agent_comment_creates_comment_parts_and_round_trips() -> Result<()> {
        let tmp = tempdir()?;
        let source = tmp.path().join("input.docx");
        let output = tmp.path().join("output.docx");
        write_test_docx(&source, TEST_DOCUMENT_XML_SIMPLE, None)?;

        let report = add_agent_comment(AddAgentCommentRequest {
            document_path: source.clone(),
            output_path: output.clone(),
            comment_text: "@Agent review this clause".to_string(),
            author: Some("Michael Wong".to_string()),
            search_text: Some("beta gamma".to_string()),
            occurrence: 1,
            paragraph_index: None,
            start_char: None,
            end_char: None,
        })?;

        assert_eq!(report.selected_text, "beta gamma");
        let scan = scan_agent_comments(&output)?;
        assert_eq!(scan.task_count, 1);
        assert_eq!(scan.tasks[0].selected_text, "beta gamma");
        assert_eq!(scan.tasks[0].author.as_deref(), Some("Michael Wong"));
        Ok(())
    }

    #[test]
    fn add_agent_comment_uses_requested_occurrence() -> Result<()> {
        let tmp = tempdir()?;
        let source = tmp.path().join("input.docx");
        let output = tmp.path().join("output.docx");
        write_test_docx(&source, TEST_DOCUMENT_XML_REPEATED, None)?;

        let report = add_agent_comment(AddAgentCommentRequest {
            document_path: source.clone(),
            output_path: output.clone(),
            comment_text: "@Agent act on the second match".to_string(),
            author: None,
            search_text: Some("repeat".to_string()),
            occurrence: 2,
            paragraph_index: None,
            start_char: None,
            end_char: None,
        })?;

        assert_eq!(report.start_char, 7);
        let scan = scan_agent_comments(&output)?;
        assert_eq!(scan.tasks[0].selected_text, "repeat");
        Ok(())
    }

    #[test]
    fn resolve_agent_comment_context_returns_window() -> Result<()> {
        let tmp = tempdir()?;
        let doc_path = tmp.path().join("sample.docx");
        write_test_docx(
            &doc_path,
            TEST_DOCUMENT_XML_WITH_COMMENT,
            Some(TEST_COMMENTS_XML),
        )?;

        let context = resolve_agent_comment_context(&doc_path, "comment-0", 1)?;
        assert_eq!(context.context_paragraphs.len(), 3);
        assert_eq!(context.task.selected_text, "targeted text");
        Ok(())
    }

    #[test]
    fn scan_agent_comments_preserves_multiline_selection() -> Result<()> {
        let tmp = tempdir()?;
        let doc_path = tmp.path().join("multiline.docx");
        write_test_docx(
            &doc_path,
            TEST_DOCUMENT_XML_WITH_MULTIPARAGRAPH_COMMENT,
            Some(TEST_COMMENTS_XML),
        )?;

        let scan = scan_agent_comments(&doc_path)?;
        assert_eq!(scan.task_count, 1);
        assert_eq!(scan.tasks[0].selected_text, "first line\nsecond line");
        Ok(())
    }

    #[test]
    fn explicit_range_validation_rejects_bad_offsets() -> Result<()> {
        let tmp = tempdir()?;
        let source = tmp.path().join("input.docx");
        write_test_docx(&source, TEST_DOCUMENT_XML_SIMPLE, None)?;

        let err = add_agent_comment(AddAgentCommentRequest {
            document_path: source.clone(),
            output_path: tmp.path().join("out.docx"),
            comment_text: "@Agent test".to_string(),
            author: None,
            search_text: None,
            occurrence: 1,
            paragraph_index: Some(0),
            start_char: Some(7),
            end_char: Some(3),
        })
        .unwrap_err();
        assert!(
            err.to_string()
                .contains("start_char must be less than end_char")
        );
        Ok(())
    }

    fn write_test_docx(path: &Path, document_xml: &str, comments_xml: Option<&str>) -> Result<()> {
        let file = fs::File::create(path)?;
        let mut writer = ZipWriter::new(file);
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
        writer.start_file("[Content_Types].xml", options)?;
        writer.write_all(content_types_xml(comments_xml.is_some()).as_bytes())?;
        writer.start_file("_rels/.rels", options)?;
        writer.write_all(ROOT_RELS_XML.as_bytes())?;
        writer.start_file("word/document.xml", options)?;
        writer.write_all(document_xml.as_bytes())?;
        writer.start_file("word/_rels/document.xml.rels", options)?;
        writer.write_all(document_rels_xml(comments_xml.is_some()).as_bytes())?;
        if let Some(comments_xml) = comments_xml {
            writer.start_file("word/comments.xml", options)?;
            writer.write_all(comments_xml.as_bytes())?;
        }
        writer.finish()?;
        Ok(())
    }

    fn content_types_xml(has_comments: bool) -> String {
        let extra = if has_comments {
            format!(
                "<Override PartName=\"/word/comments.xml\" ContentType=\"{}\"/>",
                COMMENTS_CONTENT_TYPE
            )
        } else {
            String::new()
        };
        format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
            <Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
            <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
            <Default Extension=\"xml\" ContentType=\"application/xml\"/>\
            <Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\
            {extra}\
            </Types>"
        )
    }

    fn document_rels_xml(has_comments: bool) -> String {
        let extra = if has_comments {
            format!(
                "<Relationship Id=\"rId2\" Type=\"{}\" Target=\"comments.xml\"/>",
                COMMENTS_REL_TYPE
            )
        } else {
            String::new()
        };
        format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
            <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
            <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\
            {extra}\
            </Relationships>"
        )
    }

    const ROOT_RELS_XML: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
        <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
        <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>\
        </Relationships>";

    const TEST_DOCUMENT_XML_SIMPLE: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
        <w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\
        <w:body>\
        <w:p><w:r><w:t>alpha beta gamma delta</w:t></w:r></w:p>\
        </w:body></w:document>";

    const TEST_DOCUMENT_XML_REPEATED: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
        <w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\
        <w:body>\
        <w:p><w:r><w:t>repeat repeat end</w:t></w:r></w:p>\
        </w:body></w:document>";

    const TEST_DOCUMENT_XML_WITH_COMMENT: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
        <w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\
        <w:body>\
        <w:p><w:r><w:t>Intro paragraph.</w:t></w:r></w:p>\
        <w:p>\
        <w:r><w:t>Before </w:t></w:r>\
        <w:commentRangeStart w:id=\"0\"/>\
        <w:r><w:t>targeted </w:t></w:r>\
        <w:r><w:t>text</w:t></w:r>\
        <w:commentRangeEnd w:id=\"0\"/>\
        <w:r><w:commentReference w:id=\"0\"/></w:r>\
        <w:r><w:t> after.</w:t></w:r>\
        </w:p>\
        <w:p><w:r><w:t>Closing paragraph.</w:t></w:r></w:p>\
        </w:body></w:document>";

    const TEST_DOCUMENT_XML_WITH_MULTIPARAGRAPH_COMMENT: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
        <w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\
        <w:body>\
        <w:p>\
        <w:commentRangeStart w:id=\"0\"/>\
        <w:r><w:t>first line</w:t></w:r>\
        </w:p>\
        <w:p>\
        <w:r><w:t>second line</w:t></w:r>\
        <w:commentRangeEnd w:id=\"0\"/>\
        <w:r><w:commentReference w:id=\"0\"/></w:r>\
        </w:p>\
        </w:body></w:document>";

    const TEST_COMMENTS_XML: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
        <w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\
        <w:comment w:id=\"0\" w:author=\"Reviewer\"><w:p><w:r><w:t>@Agent rewrite this section</w:t></w:r></w:p></w:comment>\
        </w:comments>";
}
