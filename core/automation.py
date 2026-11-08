"""Browser-side automation for Kara ASP.NET WebForms yearly shift assignment.
The user must sign in manually. No credentials/cookies are read or exported.
"""
from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path


def build_employee_plan(employee, mapping):
    days = {}
    errors = []
    for day, raw in employee["days"].items():
        key = (raw or "").strip() or "خالی"
        value = mapping.get(key, "ERROR")
        if value == "ERROR":
            errors.append(f"روز {day}: کد «{key}» نگاشت نشده است")
        else:
            days[int(day)] = value
    return days, errors


class OperationLog:
    def __init__(self, path=None):
        self.path = Path(path or (Path.home() / "KaraShiftManager" / "operation-log.csv"))
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if not self.path.exists():
            with self.path.open("w", newline="", encoding="utf-8-sig") as f:
                csv.writer(f).writerow(["timestamp", "national_id", "name", "shift_type", "sheet", "excel_row", "status", "detail"])

    def append(self, employee, status, detail=""):
        with self.path.open("a", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerow([datetime.now().isoformat(timespec="seconds"), employee.get("nid", ""),
                employee.get("first", "") + " " + employee.get("last", ""), employee.get("shift", ""),
                employee.get("sheet", ""), employee.get("row", ""), status, detail])

    def completed_keys(self):
        done = set()
        try:
            with self.path.open(newline="", encoding="utf-8-sig") as f:
                for row in list(csv.DictReader(f)):
                    if row.get("status") == "موفق":
                        done.add((row.get("national_id"), row.get("sheet"), row.get("excel_row")))
        except Exception:
            pass
        return done
