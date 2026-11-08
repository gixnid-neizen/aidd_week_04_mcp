import json
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

from make_report import build_workbook, deadline_status, fetch_tasks, write_report


def task(task_id, done=False, priority="normal", due=None, title=None):
    return {
        "id": task_id,
        "title": title or f"Task {task_id}",
        "done": done,
        "priority": priority,
        "due": due,
        "created_at": "2026-11-01T10:00:00",
    }


class MakeReportTest(unittest.TestCase):
    def test_deadline_status_definitions(self):
        today = date(2026, 11, 8)
        self.assertEqual(deadline_status(task(1, due=None), today), "No deadline")
        self.assertEqual(deadline_status(task(2, due="2026-11-07"), today), "Overdue")
        self.assertEqual(deadline_status(task(3, due="2026-11-08"), today), "Due today")
        self.assertEqual(deadline_status(task(4, due="2026-11-09"), today), "Upcoming")
        self.assertEqual(deadline_status(task(5, done=True, due="2026-11-07"), today), "Completed")

    def test_workbook_metrics_and_details(self):
        tasks = [
            task(1, priority="high", due="2026-11-07"),
            task(2, done=True, priority="high", due="2026-11-07"),
            task(3, due=None),
        ]
        workbook = build_workbook(tasks, date(2026, 11, 8))
        summary = workbook["Summary"]
        self.assertEqual(summary["B3"].value, 3)
        self.assertEqual(summary["B4"].value, 1)
        self.assertEqual(summary["B5"].value, 2)
        self.assertEqual(summary["B7"].value, 1)
        self.assertEqual(summary["B8"].value, 1)
        self.assertEqual(workbook["Tasks"]["F2"].value, "Overdue")
        self.assertEqual(workbook["Tasks"]["F3"].value, "Completed")
        self.assertEqual(workbook["Tasks"].max_row, 4)

    def test_empty_workbook_has_zero_completion_rate(self):
        workbook = build_workbook([], date(2026, 11, 8))
        self.assertEqual(workbook["Summary"]["B6"].value, 0)

    def test_fetch_tasks_validates_response_shape(self):
        response = unittest.mock.Mock()
        response.__enter__ = unittest.mock.Mock(return_value=response)
        response.__exit__ = unittest.mock.Mock(return_value=False)
        response.read = unittest.mock.Mock(return_value=json.dumps([task(1)]).encode())
        with patch("make_report.urllib.request.urlopen", return_value=response):
            self.assertEqual(fetch_tasks("http://todo.test")[0]["id"], 1)

        response.read = unittest.mock.Mock(return_value=b"{}")
        with patch("make_report.urllib.request.urlopen", return_value=response):
            with self.assertRaisesRegex(ValueError, "JSON array"):
                fetch_tasks("http://todo.test")

    def test_write_report_creates_valid_xlsx(self):
        tasks = [task(1, priority="high", due="2026-11-09")]
        response = unittest.mock.Mock()
        response.__enter__ = unittest.mock.Mock(return_value=response)
        response.__exit__ = unittest.mock.Mock(return_value=False)
        response.read = unittest.mock.Mock(return_value=json.dumps(tasks).encode())
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "nested" / "report.xlsx"
            with patch("make_report.urllib.request.urlopen", return_value=response):
                self.assertEqual(write_report("http://todo.test", output, date(2026, 11, 8)), 1)
            workbook = load_workbook(output, read_only=True)
            self.assertEqual(workbook.sheetnames, ["Summary", "Tasks"])
            self.assertEqual(workbook["Tasks"]["B2"].value, "Task 1")
            workbook.close()


if __name__ == "__main__":
    unittest.main()
