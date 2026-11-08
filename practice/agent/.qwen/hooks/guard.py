#!/usr/bin/env python3
"""Qwen PreToolUse hook: deny network-capable shell commands."""
import json
import re
import sys

data = json.load(sys.stdin)
tool = data.get("tool_name", "")
args = data.get("tool_input", {})


def decide(decision: str, reason: str) -> None:
    print(json.dumps({
        "hookSpecificOutput": {
            "hookEventName": "PreToolUse",
            "permissionDecision": decision,
            "permissionDecisionReason": reason,
        }
    }, ensure_ascii=False))
    sys.exit(0)


command = args.get("command", "") if isinstance(args, dict) else ""
if tool == "run_shell_command" and isinstance(command, str) and re.search(
    r"\b(curl|wget|nc|ncat)\b", command, flags=re.IGNORECASE
):
    decide("deny", "Сетевые команды запрещены.")

decide("allow", "ok")
