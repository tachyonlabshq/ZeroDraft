# ZeroDraft `@Agent` Comment Workflow

## Goal

Allow humans to use native Microsoft Word comments as precise task assignments for AI agents.

## How it works

1. The user highlights text in Word.
2. The user creates a comment that begins with `@Agent`.
3. ZeroDraft scans the DOCX and extracts:
   - the selected text range
   - the raw comment text
   - the normalized instruction
   - the paragraph span containing the selection

## Example

Highlighted text:

`Limitation of Liability`

Word comment:

`@Agent make this narrower and align it with the indemnity section`

ZeroDraft returns a task payload with both the highlighted target and the instruction.

## Current v1 scope

- classic DOCX comments
- case-insensitive `@Agent` detection
- exact highlighted-text recovery from comment range markers
- paragraph-window context resolution
- programmatic insertion of classic comments for targeted follow-up

## Current v1 limits

- insertion targets one paragraph at a time
- insertion expects commentable text runs and may reject unusually complex run structures
- threaded Word comments are not yet implemented
- tracked-changes-aware editing is not yet implemented

## Recommended agent loop

1. `inspect_document`
2. `scan_agent_comments`
3. `resolve_agent_comment_context`
4. `plan_agent_comment`
5. draft a change
6. optionally write back a new targeted comment with `add_agent_comment`
