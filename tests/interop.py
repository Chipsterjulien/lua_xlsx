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
    wb = load_workbook(path, data_only=False)
    assert wb.epoch == CALENDAR_MAC_1904
    assert wb.sheetnames == ["Résumé"]
    ws = wb["Résumé"]

    assert ws["A2"].value == "Élise & <test>"
    assert ws["B2"].value == 12.5
    assert ws["C2"].value == datetime(2024, 3, 15, 13, 45, 30)
    assert ws["D4"].value == "=SUM(B2:B3)"

    assert ws.freeze_panes == "B2"
    assert ws.auto_filter.ref == "A1:D3"
    assert ws.column_dimensions["A"].width == 18.0
    assert ws.column_dimensions["B"].width == 14.0
    assert ws.row_dimensions[1].height == 24.0

    assert ws["A1"].font.bold is True
    assert ws["A1"].font.color.type == "rgb"
    assert ws["A1"].font.color.rgb == "FF112233"
    assert ws["A1"].fill.fill_type == "solid"
    assert ws["A1"].fill.fgColor.rgb == "FFD9EAF7"
    assert ws["A1"].alignment.horizontal == "center"
    assert ws["A1"].alignment.vertical == "center"
    assert ws["A1"].alignment.wrap_text is True
    assert ws["B2"].number_format == '#,##0.00 "€"'
    assert ws["C2"].font.bold is True
    assert ws["C2"].number_format == "yyyy-mm-dd hh:mm:ss"

    data_wb = load_workbook(path, data_only=True)
    data_ws = data_wb["Résumé"]
    assert data_ws["C2"].value == datetime(2024, 3, 15, 13, 45, 30)
    assert data_ws["D4"].value is None

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
