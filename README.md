# ZeroDraft

ZeroDraft is a Rust-native Word and DOCX skill for AI agents. It runs as both a local CLI and an MCP server, and it is designed to fit the same Zero-family workflow as ZeroCell while focusing on document text, comment ranges, and agent-targeted follow-up inside Microsoft Word files.

## Core capabilities

- Inspect a DOCX package for paragraphs, tables, and comment inventory.
- Extract visible paragraph text for bounded agent context.
- Scan Word comments tagged with `@Agent` and return:
  - the highlighted text range
  - the raw comment text
  - the normalized instruction
  - the paragraph window where the comment applies
- Resolve surrounding paragraph context for a scanned `@Agent` task.
- Dry-run a comment insertion plan before mutating a document.
- Insert classic Word comments programmatically for targeted follow-up.
  - supports multi-run selections, including visible tabs and line breaks within a run
- Convert legacy `.doc` files to `.docx` through LibreOffice headless mode.
- Run as `mcp-stdio` for OpenCode and other MCP-capable hosts.

## The `@Agent` workflow

ZeroDraft’s primary differentiator is comment-native agent task extraction.

1. A human opens a Word document.
2. They highlight a passage.
3. They create a Word comment beginning with `@Agent`.
4. ZeroDraft scans the document and returns both:
   - the highlighted text as the task target
   - the comment body as the instruction

This makes the document itself the task surface instead of a separate prompt log.

See [docs/AGENT_COMMENT_WORKFLOW.md](/Users/michaelwong/Developer/ZeroDraft/docs/AGENT_COMMENT_WORKFLOW.md).

## Commands

- `inspect-document`
- `extract-text`
- `scan-agent-comments`
- `resolve-agent-comment-context`
- `plan-agent-comment`
- `add-agent-comment`
- `replace-range-text`
- `convert-to-docx`
- `doctor`
- `init`
- `schema-info`
- `skill-api-contract`
- `mcp-stdio`

## Example usage

Inspect a document:

```bash
cargo run -- inspect-document ./contract.docx --pretty
```

Extract `@Agent` tasks from Word comments:

```bash
cargo run -- scan-agent-comments ./contract.docx --pretty
```

Resolve nearby context for a task:

```bash
cargo run -- resolve-agent-comment-context ./contract.docx comment-0 --window-radius 2 --pretty
```

Plan a targeted comment without mutating the document:

```bash
cargo run -- plan-agent-comment ./contract.docx \
  --search-text "Limitation of Liability" \
  --comment-text "@Agent tighten this clause" \
  --pretty
```

Create a targeted comment on matching text:

```bash
cargo run -- add-agent-comment ./contract.docx ./contract.commented.docx \
  --search-text "Limitation of Liability" \
  --comment-text "@Agent tighten this clause" \
  --author "ZeroDraft" \
  --pretty
```

Replace a targeted text range into a new DOCX:

```bash
cargo run -- replace-range-text ./contract.docx ./contract.revised.docx \
  --search-text "Limitation of Liability" \
  --replacement-text "Mutual Limitation of Liability" \
  --pretty
```

Convert a legacy `.doc` file:

```bash
cargo run -- convert-to-docx ./legacy.doc ./legacy.docx --pretty
```

## MCP setup

The repository includes [mcp.json](/Users/michaelwong/Developer/ZeroDraft/mcp.json) for an in-repo OpenCode registration example. The `init` command can also merge a local ZeroDraft entry into an existing `opencode.json`.

## Binaries

Packaged binaries live under `bin/`. Local builds from this machine currently target macOS first; the GitHub Actions release workflow is configured to produce additional macOS and Windows artifacts.

## Platform Bundles

The repository ships an install-ready bundle workflow at [.github/workflows/platform-bundles.yml](/Users/michaelwong/Developer/ZeroDraft/.github/workflows/platform-bundles.yml). It builds one self-contained zip per platform:

- `ZeroDraft-macos-arm64-<version>.zip`
- `ZeroDraft-macos-x64-<version>.zip`
- `ZeroDraft-windows-x64-<version>.zip`
- `ZeroDraft-windows-arm64-<version>.zip`

Each zip extracts to exactly one top-level folder:

```text
ZeroDraft/
  README.md
  SKILL.md
  mcp.json
  bin/
    zerodraft or zerodraft.exe
```

The included `mcp.json` already points at the bundled local binary via a relative `./bin/...` path and passes `mcp-stdio`, so users can download the zip, extract it, and drop the resulting `ZeroDraft/` folder directly into their skills directory.

The workflow uploads per-platform zips, manifests, and checksum files as artifacts, then generates aggregate bundle metadata and SHA256 sums. On `v*` tags it publishes all of those files to the GitHub Release.

## Validation status

Current local validation in this repo:

- `cargo test`
- `cargo build --release`

Planned release gates are codified in CI:

- `cargo fmt --check`
- `cargo check`
- `cargo clippy --all-targets --all-features -- -D warnings`
- `cargo audit`
- `cargo deny check`

## Repo structure

- [src](/Users/michaelwong/Developer/ZeroDraft/src)
- [docs](/Users/michaelwong/Developer/ZeroDraft/docs)
- [bin](/Users/michaelwong/Developer/ZeroDraft/bin)
- [ROADMAP.md](/Users/michaelwong/Developer/ZeroDraft/ROADMAP.md)
- [SKILL.md](/Users/michaelwong/Developer/ZeroDraft/SKILL.md)
