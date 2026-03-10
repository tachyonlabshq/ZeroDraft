# ZeroDraft Roadmap

ZeroDraft is a Rust-native Word and DOCX skill for AI agents. The first production target is a single binary that can run as both a CLI and an MCP server, with an OpenCode-friendly distribution bundle and a public packaging repository containing prebuilt binaries.

## Phase 0: Product Definition and Repo Baseline

- [x] Define ZeroDraft as part of the Zero- office AI family.
- [x] Choose the initial architecture:
  - One self-contained Rust repo for source, tests, docs, release automation, and packaged binaries.
  - Distribution assets live in-repo under stable paths such as `bin/`, `SKILL.md`, `README.md`, and `mcp.json`.
  - Native `.docx` support first, with `.doc` handled through an optional conversion path when available.
- [x] Freeze naming and contract surface:
  - Binary name: `zerodraft`
  - Core MCP namespace: `zerodraft`
  - Stable tool contract version for CLI and MCP payloads.

## Phase 1: Core OOXML Document Engine

- [x] Scaffold the Rust workspace and release profile.
- [x] Implement DOCX archive reading with deterministic ZIP and OOXML validation:
  - detect required parts such as `[Content_Types].xml`, `word/document.xml`, and relationships
  - return actionable diagnostics for malformed or unsupported packages
  - preserve unknown parts while mutating targeted document/comment content
- [x] Implement document text extraction and structure indexing:
  - paragraphs
  - runs
  - tables and table cells
  - bookmarks and simple section metadata
- [x] Build a stable internal address model for text ranges:
  - paragraph index
  - run index
  - character offsets within merged visible text
  - comment anchor metadata that can be mapped back into OOXML boundaries

## Phase 2: Agent-Oriented Comment Workflow

- [x] Implement comment scanning for classic Word comments:
  - parse `comments.xml`
  - detect `@Agent` or `@agent`
  - capture author, comment id, and instruction text
  - capture the highlighted anchor range from comment markers in `document.xml`
- [x] Implement a range-aware `@Agent` task payload:
  - document path
  - deterministic task id
  - anchor paragraph/run metadata
  - selected text snippet
  - raw comment text
  - normalized instruction
  - optional surrounding context window
- [x] Implement comment insertion for follow-up tasks:
  - accept explicit target text or indexed range
  - apply `w:commentRangeStart`, `w:commentRangeEnd`, and `w:commentReference`
  - create comment parts and relationships if missing
  - preserve non-targeted document content
- [x] Validate edge cases:
  - repeated text matches
  - split runs with different formatting
  - multiple text nodes, tabs, and line breaks within a run
  - empty selections
  - existing comments and id collisions
  - multi-paragraph selections

## Phase 3: Editing and Transformation Tools

- [x] Implement safe document inspection and export tools:
  - extract plain text
  - inspect structure
  - read comment tasks
  - read a bounded range around a target selection
- [x] Implement controlled write operations:
  - replace targeted text by explicit range
  - append paragraphs
  - add `@Agent` comments without rewriting unrelated XML
  - optionally convert `.doc` to `.docx` through LibreOffice or another local converter when present
- [x] Add plan/validate/apply style operations where mutation risk is non-trivial:
  - dry-run comment insertion planning implemented
  - comment insertion validation implemented through shared target resolution
  - output path enforcement for non-destructive workflows implemented

## Phase 4: MCP and Agent Compatibility

- [x] Expose the core features through CLI and MCP stdio using one binary.
- [x] Define concise, agent-discoverable tools with stable JSON schemas:
  - `inspect_document`
  - `extract_text`
  - `extract_range`
  - `scan_agent_comments`
  - `add_agent_comment`
  - `replace_range_text`
  - `convert_to_docx`
  - `schema_info`
  - `skill_api_contract`
  - `mcp_stdio`
- [x] Make the responses bounded and structured for agent loops:
  - predictable statuses
  - structured content with machine-usable metadata
  - actionable error categories and suggested next steps

## Phase 5: Testing, Debugging, and Hardening

- [x] Build a representative fixture corpus:
  - minimal DOCX
  - styled paragraphs
  - tables
  - existing comments
  - documents with split runs
  - malformed archives
- [x] Add unit and integration-style tests for:
  - text extraction
  - range mapping
  - `@Agent` comment detection
  - comment insertion round-trips
  - MCP tool invocation behavior
- [x] Run iterative debugging until the core flows are validated:
  - clean documents
  - heavily formatted documents
  - duplicate text anchors
  - comment insertion into already commented regions
- [x] Add regression coverage for every bug found during iteration.

## Phase 6: Packaging and Distribution

- [x] Create release automation for:
  - macOS arm64
  - macOS x64
  - Windows x64
  - Windows arm64
  - install-ready per-platform zip bundles with manifests and checksums
  - Ubuntu-hosted Windows builds install and shim LLVM tools required by `cargo-xwin` and `cc-rs`
  - successful `main` builds publish a rolling prerelease, while `v*` tags publish versioned releases
- [x] Publish a self-contained in-repo bundle layout:
  - `bin/`
  - `SKILL.md`
  - `README.md`
  - `mcp.json`
- [x] Ensure the main repository supports direct OpenCode installation and generic MCP usage without a second packaging repo.

## Phase 7: Security, Operations, and Release

- [x] Run quality gates:
  - `cargo fmt --check`
  - `cargo check`
  - `cargo test`
  - `cargo clippy --all-targets --all-features -- -D warnings`
- [x] Run supply-chain and security checks:
  - `cargo audit`
  - `cargo deny check`
  - inspect archive write paths and XML handling for traversal or injection issues
- [x] Write operator docs:
  - setup
  - local MCP registration
  - `.doc` conversion prerequisites
  - `@Agent` workflow semantics
- [x] Create the GitHub repo under `tachyonlabshq`, commit, push, and verify the public history.

## Phase 8: Post-v1 Expansion

- [ ] Add richer DOCX editing primitives:
  - tables
  - headers and footers
  - tracked changes aware mutation
  - style-aware paragraph insertion
- [ ] Evaluate higher-fidelity rendering or conversion support using additional Rust libraries.
- [ ] Add broader agent ecosystem packaging beyond OpenCode where useful.
