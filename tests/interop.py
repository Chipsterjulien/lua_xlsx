#!/usr/bin/env python3
from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils.datetime import CALENDAR_MAC_1904


def prepare(path: Path) -> None:
    wb = Workbook()
    wb.epoch = CALENDAR_MAC_1904
    ws = wb.active
    ws.title = "Entrée"
    ws.append(["Nom", "Actif", "Date", "Formule"])
    ws.append(["Zoé", False, datetime(2025, 1, 2, 3, 4, 5), "=SUM(1,2)"])
    ws.append(["Autre", True, datetime(2025, 1, 3, 4, 5, 6), 7])
    ws["A2"].hyperlink = "https://example.org/source?a=1&b=2"
    ws["A2"].hyperlink.tooltip = "Source externe"
    ws["B2"].font = Font(
        name="DejaVu Sans", size=13, bold=True, italic=True,
        underline="single", strike=True, color="FF123456",
    )
    ws["B2"].fill = PatternFill(fill_type="solid", fgColor="FFABCDEF")
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["B2"].border = Border(
        left=Side(style="thin", color="FFFF0000"),
        right=Side(style="double"),
        top=Side(style="dashed", color="FF00AA00"),
        bottom=Side(style="dotted"),
    )
    ws.column_dimensions["A"].width = 20
    ws.row_dimensions[1].height = 26
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = "A1:D3"
    ws.merge_cells("A4:C4")
    ws["A4"] = "Fusion openpyxl"
    ws.sheet_properties.tabColor = "5B9BD5"
    ws["C2"].comment = Comment("Commentaire openpyxl", "Interop")
    ws.row_dimensions[5].hidden = True
    ws.column_dimensions["E"].hidden = True
    dv_list = DataValidation(
        type="list", formula1='"Oui,Non,En attente"', allow_blank=True,
        showInputMessage=True, showErrorMessage=True, promptTitle="Choix",
        prompt="Sélectionner", errorTitle="Erreur", error="Valeur invalide",
    )
    ws.add_data_validation(dv_list)
    dv_list.add("A2:A10")
    dv_whole = DataValidation(type="whole", operator="between", formula1="0", formula2="100")
    ws.add_data_validation(dv_whole)
    dv_whole.add("B2:B10")
    ws.conditional_formatting.add(
        "B2:B10",
        CellIsRule(
            operator="lessThan", formula=["0"],
            font=Font(color="FF9C0006"),
            fill=PatternFill(fill_type="solid", fgColor="FFFFC7CE"),
        ),
    )
    target = wb.create_sheet("Cible")
    target["A1"] = "destination"
    hidden = wb.create_sheet("Masquée")
    hidden.sheet_state = "hidden"
    very_hidden = wb.create_sheet("Très masquée")
    very_hidden.sheet_state = "veryHidden"
    wb.active = 0
    wb.defined_names.add(DefinedName("ZoneNoms", attr_text="'Entrée'!$A$2:$A$10"))
    wb.defined_names.add(DefinedName("ValeurLocale", attr_text="'Entrée'!$B$2", localSheetId=0, hidden=True))
    wb.save(path)
    print(f"Fixture openpyxl créée : {path}")


def check(path: Path) -> None:
    wb = load_workbook(path, data_only=False)
    assert wb.epoch == CALENDAR_MAC_1904
    assert wb.sheetnames == ["Résumé", "Cible", "Masquée", "Très masquée"]
    ws = wb["Résumé"]

    assert ws["A1"].value == "Rapport 1.3"
    assert str(next(iter(ws.merged_cells.ranges))) == "A1:D1"
    assert ws["A3"].value == "Élise & <test>"
    assert ws["B3"].value == 12.5
    assert ws["C3"].value == datetime(2024, 3, 15, 13, 45, 30)
    assert ws["D5"].value == "=SUM(B3:B4)"

    assert ws.freeze_panes == "B3"
    assert ws.auto_filter.ref == "A2:D4"
    assert ws.column_dimensions["A"].width == 18.0
    assert ws.column_dimensions["B"].width == 14.0
    assert ws.row_dimensions[1].height == 28.0

    assert ws["A1"].font.bold is True
    assert ws["A1"].font.italic is True
    assert ws["A1"].font.name == "Liberation Sans"
    assert ws["A1"].font.sz == 14.0
    assert ws["A1"].font.underline == "double"
    assert ws["A1"].font.strike is True
    assert ws["A1"].font.color.type == "rgb"
    assert ws["A1"].font.color.rgb == "FF112233"
    assert ws["A1"].fill.fill_type == "solid"
    assert ws["A1"].fill.fgColor.rgb == "FFD9EAF7"
    assert ws["A1"].alignment.horizontal == "center"
    assert ws["A1"].alignment.vertical == "center"
    assert ws["A1"].alignment.wrap_text is True
    assert ws["A1"].border.left.style == "thin"
    assert ws["A1"].border.left.color.rgb == "FFFF0000"
    assert ws["A1"].border.right.style == "double"
    assert ws["A1"].border.top.style == "dashed"
    assert ws["A1"].border.bottom.style == "dotted"

    assert ws["B3"].number_format == '#,##0.00 "€"'
    assert ws["C3"].font.bold is True
    assert ws["C3"].number_format == "yyyy-mm-dd hh:mm:ss"

    assert ws["D3"].hyperlink.target == "https://example.com/?a=1&b=2"
    assert ws["D3"].hyperlink.tooltip == "Ouvrir le site"
    assert ws["D4"].hyperlink.location == "'Cible'!A1"

    assert wb.active.title == "Résumé"
    assert ws.sheet_properties.tabColor.rgb == "FF4472C4"
    assert ws.row_dimensions[6].hidden is True
    assert ws.column_dimensions["E"].hidden is True
    assert ws["A4"].comment.text == "Valeur contrôlée par lua-xlsx"
    assert ws["A4"].comment.author == "Julien"
    validations = list(ws.data_validations.dataValidation)
    assert len(validations) == 2
    assert validations[0].type == "list" and str(validations[0].sqref) == "A3:A20"
    assert validations[0].formula1 == '"Oui,Non,En attente"'
    assert validations[1].type == "whole" and validations[1].operator == "between"
    assert validations[1].formula1 == "0" and validations[1].formula2 == "100"
    assert len(ws.conditional_formatting) == 1
    rules = next(iter(ws.conditional_formatting._cf_rules.values()))
    assert rules[0].type == "cellIs" and rules[0].operator == "lessThan"
    assert rules[0].dxf.fill.fgColor.rgb == "FFFFC7CE"
    assert wb["Masquée"].sheet_state == "hidden"
    assert wb["Très masquée"].sheet_state == "veryHidden"
    names = {name.name: name for name in wb.defined_names.values()}
    assert names["ZoneNoms"].attr_text == "'Résumé'!$A$3:$A$20"
    local_name = ws.defined_names["ValeurLocale"]
    assert local_name.localSheetId == 0 and local_name.hidden is True

    data_wb = load_workbook(path, data_only=True)
    data_ws = data_wb["Résumé"]
    assert data_ws["C3"].value == datetime(2024, 3, 15, 13, 45, 30)
    assert data_ws["D5"].value == 19.75

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
