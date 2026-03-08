# ZeroDraft Skill

## Purpose

Use ZeroDraft when an agent needs to inspect or manipulate Microsoft Word documents through a single compiled Rust binary.

## Primary strengths

- DOCX inspection and text extraction
- Word comment scanning for `@Agent` tasks
- Highlighted-range targeting from native Word comments
- Comment-context resolution for follow-up actions
- Optional `.doc` to `.docx` conversion
- MCP compatibility through `mcp-stdio`

## Recommended workflow

1. Run `inspect-document` on unfamiliar files.
2. Run `scan-agent-comments` to discover user-authored `@Agent` tasks.
3. Run `resolve-agent-comment-context` before drafting a targeted change.
4. Use `add-agent-comment` when the agent needs to leave a precise follow-up marker.
5. Use `convert-to-docx` when the source file is a legacy `.doc`.

## Binary

- `bin/zerodraft-macos-arm64`
- additional packaged binaries are produced by the release workflow and stored in `bin/` when available

## MCP tools

- `inspect_document`
- `extract_text`
- `scan_agent_comments`
- `resolve_agent_comment_context`
- `add_agent_comment`
- `convert_to_docx`
- `doctor_environment`
- `schema_info`
- `skill_api_contract`
