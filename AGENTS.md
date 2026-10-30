# Repository guidance

## Git approval boundary
- Do not execute Git commands unless the user explicitly approves the specific command first.
- `git status` is the only Git command pre-approved by the user. It may be run without asking.
- Approval for one Git command does not authorize any other Git command or a later repeat.
- After the user approves an exact command in chat, write that exact command alone to `.codex/.approved_git_command` immediately before running it. The repository PreToolUse hook consumes this one-time authorization. Never pre-authorize commands.
- Do not use Git through scripts, IDE integrations, hooks, or other tools to bypass this boundary.

## Engineering rules
- Keep API credentials in environment variables. Never print, log, commit, or expose their values.
- Validate and bound all user input at the HTTP boundary; return actionable client errors and generic server errors.
- Never invent presentation content as a successful fallback when both AI providers fail. Tell the user generation is unavailable.
- Keep generated PPTX files in memory; do not persist user prompts or generated files unless the user asks.
- Keep Python changes compatible with the runtime pinned in `.python-version` and `runtime.txt`.
- Do not add dependencies without a clear need and document required environment variables in README.
- Keep UI output accessible, responsive, and explicit about generation progress and failure.
- Before completing a change, inspect the diff and run relevant checks when the user asks for verification. Never run Git commands other than `git status` without explicit approval.
