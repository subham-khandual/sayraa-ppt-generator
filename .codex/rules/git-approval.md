# Git command approval

Do not run any Git command without the user's explicit approval for that exact command. The sole exception is `git status`, which is pre-approved. This applies to read-only commands, shell scripts, IDE actions, and commands invoked indirectly. Never assume approval carries over to another command. After the user approves a command in chat, place only that exact command in `.codex/.approved_git_command` immediately before running it; the PreToolUse hook consumes the one-time authorization. Never pre-authorize commands.
