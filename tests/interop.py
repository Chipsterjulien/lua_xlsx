#!/usr/bin/env python3
from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.utils.datetime import CALENDAR_MAC_1904


def prepare(path: Path) -> None:
    wb = Workbook()
    wb.epoch = CALENDAR_MAC_1904
    ws = wb.active
    ws.title = "Entrée"
    ws.append(["Nom", "Actif", "Date"])
    ws.append(["Zoé", False, datetime(2025, 1, 2, 3, 4, 5)])
    wb.save(path)
    print(f"Fixture openpyxl créée : {path}")


def check(path: Path) -> None:
    wb = load_workbook(path, data_only=True)
    assert wb.epoch == CALENDAR_MAC_1904
    assert wb.sheetnames == ["Résumé"]
    ws = wb["Résumé"]
    assert ws["A2"].value == "Élise & <test>"
    assert ws["B2"].value == 12.5
    assert ws["C2"].value == datetime(2024, 3, 15, 13, 45, 30)
    print("INTEROP OPENPYXL : PASS")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("action", choices=("prepare", "check"))
    parser.add_argument("path", type=Path)
    args = parser.parse_args()
    if args.action == "prepare":
        prepare(args.path)
    else:
        check(args.path)


if __name__ == "__main__":
    main()
