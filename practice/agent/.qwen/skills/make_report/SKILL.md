---
name: make_report
description: Create a consistent Excel report for the TODO app, including completion, priority, and deadline metrics. Use when asked for a task report, TODO summary, or Excel report from TODO tasks.
---

# TODO task report

1. Get the TODO API URL from the user's request or the `todo` MCP server configuration (default: `http://127.0.0.1:5000`).
2. From `practice/agent`, run:

   ```bash
   python .qwen/skills/make_report/scripts/make_report.py \
     --api-url http://127.0.0.1:5000 \
     --output todo_report.xlsx
   ```

3. The script fetches all tasks from `/api/tasks` and saves the workbook at the requested output path.
4. Verify that the file exists, opens as an `.xlsx`, and contains the `Summary` and `Tasks` sheets. Report any API or file errors; do not claim success if generation fails.

The summary counts total, completed, open, overdue open, undated, and high-priority open tasks, with additional totals by priority. A task is overdue only if it is still open and its due date is before the report date.
