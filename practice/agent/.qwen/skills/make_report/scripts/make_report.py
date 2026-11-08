#!/usr/bin/env python3
"""Build a deterministic Excel summary from the TODO app API."""
from __future__ import annotations

import argparse
import json
import sys
import urllib.error
import urllib.request
from datetime import date
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

PRIORITIES = ("low", "normal", "high")
PRIORITY_ORDER = {priority: index for index, priority in enumerate(PRIORITIES)}


def fetch_tasks(api_url: str, timeout: float = 5) -> list[dict[str, Any]]:
    request = urllib.request.Request(
        f"{api_url.rstrip('/')}/api/tasks",
        headers={"Accept": "application/json"},
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            tasks = json.load(response)
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError) as error:
        raise RuntimeError(f"Could not fetch TODO tasks from {api_url}: {error}") from error

    if not isinstance(tasks, list):
        raise ValueError("The TODO API response must be a JSON array")
    for index, task in enumerate(tasks):
        if not isinstance(task, dict):
            raise ValueError(f"Task at index {index} must be a JSON object")
        if not isinstance(task.get("id"), int) or isinstance(task.get("id"), bool):
            raise ValueError(f"Task at index {index} has an invalid id")
        if not isinstance(task.get("title"), str):
            raise ValueError(f"Task at index {index} has an invalid title")
        if not isinstance(task.get("done"), bool):
            raise ValueError(f"Task at index {index} has an invalid done status")
        if task.get("priority") not in PRIORITY_ORDER:
            raise ValueError(f"Task at index {index} has an invalid priority")
        due = task.get("due")
        if due is not None:
            if not isinstance(due, str):
                raise ValueError(f"Task at index {index} has an invalid due date")
            try:
                date.fromisoformat(due)
            except ValueError as error:
                raise ValueError(f"Task at index {index} has an invalid due date") from error
    return tasks


def deadline_status(task: dict[str, Any], report_date: date) -> str:
    if task["done"]:
        return "Completed"
    due = task.get("due")
    if due is None:
        return "No deadline"
    due_date = date.fromisoformat(due)
    if due_date < report_date:
        return "Overdue"
    if due_date == report_date:
        return "Due today"
    return "Upcoming"


def _style_table(sheet, header_row: int = 1) -> None:
    for cell in sheet[header_row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="355C7D")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    sheet.freeze_panes = f"A{header_row + 1}"
    last_column = get_column_letter(sheet.max_column)
    sheet.auto_filter.ref = f"A{header_row}:{last_column}{sheet.max_row}"


def build_workbook(tasks: list[dict[str, Any]], report_date: date | None = None) -> Workbook:
    report_date = report_date or date.today()
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Summary"

    overdue = [task for task in tasks if deadline_status(task, report_date) == "Overdue"]
    open_tasks = [task for task in tasks if not task["done"]]
    metrics = [
        ("Report date", report_date.isoformat()),
        ("Total tasks", len(tasks)),
        ("Completed tasks", sum(task["done"] for task in tasks)),
        ("Open tasks", len(open_tasks)),
        ("Completion rate", sum(task["done"] for task in tasks) / len(tasks) if tasks else 0),
        ("Overdue open tasks", len(overdue)),
        ("Tasks without a deadline", sum(task.get("due") is None for task in tasks)),
        ("Open high-priority tasks", sum(
            not task["done"] and task["priority"] == "high" for task in tasks
        )),
    ]
    summary.append(["Metric", "Value"])
    for row in metrics:
        summary.append(list(row))
    summary["B6"].number_format = "0.0%"

    summary.append([])
    summary.append(["Priority", "Total", "Completed", "Open", "Overdue open"])
    for priority in PRIORITIES:
        group = [task for task in tasks if task["priority"] == priority]
        summary.append([
            priority,
            len(group),
            sum(task["done"] for task in group),
            sum(not task["done"] for task in group),
            sum(task in overdue for task in group),
        ])
    _style_table(summary, 1)
    _style_table(summary, 11)
    for column, width in {"A": 30, "B": 20, "C": 16, "D": 12, "E": 18}.items():
        summary.column_dimensions[column].width = width

    detail = workbook.create_sheet("Tasks")
    detail.append(["ID", "Title", "Status", "Priority", "Due date", "Deadline status", "Created at"])
    ordered = sorted(
        tasks,
        key=lambda task: (
            task.get("due") is None,
            task.get("due") or "",
            task["id"],
        ),
    )
    for task in ordered:
        detail.append([
            task["id"],
            task["title"],
            "Completed" if task["done"] else "Open",
            task["priority"],
            task.get("due") or "",
            deadline_status(task, report_date),
            task.get("created_at", ""),
        ])
    _style_table(detail)
    for index, width in enumerate((10, 48, 14, 14, 16, 20, 24), start=1):
        detail.column_dimensions[get_column_letter(index)].width = width
    for row in detail.iter_rows(min_row=2):
        row[1].alignment = Alignment(vertical="top", wrap_text=True)
    return workbook


def write_report(api_url: str, output_path: str | Path, report_date: date | None = None) -> int:
    tasks = fetch_tasks(api_url)
    target = Path(output_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    build_workbook(tasks, report_date).save(target)
    return len(tasks)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Create an Excel report from TODO tasks")
    parser.add_argument("--api-url", default="http://127.0.0.1:5000")
    parser.add_argument("--output", default="todo_report.xlsx")
    args = parser.parse_args(argv)
    try:
        count = write_report(args.api_url, args.output)
    except (RuntimeError, ValueError, OSError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1
    print(f"Created {args.output} with {count} tasks")
    return 0


if __name__ == "__main__":
    sys.exit(main())
