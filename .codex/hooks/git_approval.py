"""Block Git commands unless they are git status or one-shot approved verbatim."""

import json
import os
import re
import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[2]
APPROVAL_FILE = REPO_ROOT / ".codex" / ".approved_git_command"
GIT_TOKEN = re.compile(r"(?i)(?<![\w.-])git(?:\.exe)?(?![\w.-])")
STATUS_ONLY = re.compile(r"(?is)^\s*git(?:\.exe)?\s+status(?:\s+[^;&|`\n]+)?\s*$")


def block(reason):
    print(json.dumps({
        "hookSpecificOutput": {
            "hookEventName": "PreToolUse",
            "permissionDecision": "deny",
            "permissionDecisionReason": reason,
        }
    }))


def main():
    try:
        event = json.load(sys.stdin)
        command = event.get("tool_input", {}).get("command", "")
        if not isinstance(command, str) or not GIT_TOKEN.search(command):
            return

        if STATUS_ONLY.fullmatch(command):
            return

        try:
            approved = APPROVAL_FILE.read_text(encoding="utf-8").rstrip("\r\n")
        except OSError:
            approved = ""

        if approved and command == approved:
            try:
                APPROVAL_FILE.unlink()
            except OSError:
                block("Could not consume the one-time Git approval; command denied. Re-approve and retry.")
                return
            return

        block("Git command denied. Only git status is pre-approved. Get explicit user approval for this exact command, then place that exact command in .codex/.approved_git_command for one-time use.")
    except Exception:
        block("Git approval hook failed closed; no Git command may run.")


if __name__ == "__main__":
    main()
