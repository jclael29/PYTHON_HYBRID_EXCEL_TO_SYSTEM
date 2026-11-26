import os
import sys
import tempfile
import types
import unittest
from unittest.mock import MagicMock, patch

from openpyxl import Workbook
from openpyxl.utils import column_index_from_string

import config
import config_form9
from services.excel_utils import norm_lrn, round_nearest_whole, unique_name_lrns
from services.form9_engine import Form9Engine, _recalculate_workbooks
from services.sf10_engine import SF10Engine


def _set_cell(ws, row, column, value):
    ws.cell(row=row, column=column_index_from_string(column), value=value)


def _make_workbook(path, students, grade, is_grade_11):
    workbook = Workbook()
    master = workbook.active
    master.title = "MASTER SHEET"
    master_row = config.MASTER_START_ROW
    for student in students:
        _set_cell(master, master_row, "B", student["source_name"])
        _set_cell(master, master_row, "C", student["lrn"])
        if is_grade_11:
            _set_cell(master, master_row, "G", "Section 11")
        else:
            for key, column in config.MASTER_COLS.items():
                if key == "name":
                    value = student["output_name"]
                elif key == "lrn":
                    value = student["lrn"]
                elif key == "sex":
                    value = student["sex"]
                elif key == "last":
                    value = "Student"
                elif key == "first":
                    value = student["output_name"]
                elif key == "middle":
                    value = ""
                elif key == "section":
                    value = "Section 12"
                elif key == "school_year":
                    value = "2026-2027"
                else:
                    value = None
                _set_cell(master, master_row, column, value)
        master_row += 1

    if is_grade_11:
        for sheet_config in config.G11_SHEETS.values():
            ws = workbook.create_sheet(sheet_config["name"])
            for row_number, student in enumerate(students, config.G11_START_ROW):
                _set_cell(ws, row_number, "B", student["source_name"])
                _set_cell(ws, row_number, "C", grade)
                _set_cell(ws, row_number, "D", grade + 1)
    else:
        for sheet_name in config.G12_SHEETS:
            ws = workbook.create_sheet(sheet_name)
            for row_number, student in enumerate(students, config.G12_START_ROW):
                _set_cell(ws, row_number, "B", student["source_name"])
                _set_cell(ws, row_number, "C", grade)
    workbook.save(path)


def _make_form9_workbook(path, student, grade, is_grade_11):
    workbook = Workbook()
    master = workbook.active
    master.title = "MASTER SHEET"
    row = config_form9.MASTER_START_ROW
    if is_grade_11:
        _set_cell(master, row, "B", student["source_name"])
    _set_cell(master, row, "C", student["lrn"])
    if is_grade_11:
        for sheet_config in config_form9.G11_SHEETS.values():
            ws = workbook.create_sheet(sheet_config["name"])
            grade_row = sheet_config["start_row"]
            _set_cell(ws, grade_row, "B", student["source_name"])
            _set_cell(ws, grade_row, "E", grade)
    else:
        _set_cell(master, row, "H", student["last"])
        _set_cell(master, row, "I", student["first"])
        _set_cell(master, row, "J", student.get("middle"))
        _set_cell(master, row, "D", student["sex"])
        _set_cell(master, row, "M", "Previous School")
        _set_cell(master, row, "N", "2026-2027")
        for sheet_name, sheet_config in config_form9.G12_SHEETS.items():
            ws = workbook.create_sheet(sheet_name)
            grade_row = sheet_config["start_row"]
            _set_cell(ws, grade_row, "B", student["source_name"])
            _set_cell(ws, grade_row, "C", grade)
    workbook.save(path)


class MultiFileOutputTests(unittest.TestCase):
    def test_form9_recalculates_saved_workbooks_with_excel(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = os.path.join(temp_dir, "student.xlsx")
            Workbook().save(path)
            workbook = MagicMock()
            excel = MagicMock()
            excel.Workbooks.Open.return_value = workbook
            client = types.ModuleType("win32com.client")
            client.DispatchEx = MagicMock(return_value=excel)
            win32com = types.ModuleType("win32com")
            win32com.client = client

            with patch.dict(sys.modules, {"win32com": win32com, "win32com.client": client}):
                _recalculate_workbooks([path])

            client.DispatchEx.assert_called_once_with("Excel.Application")
            excel.Workbooks.Open.assert_called_once_with(
                os.path.abspath(path), UpdateLinks=0, ReadOnly=False
            )
            excel.CalculateFullRebuild.assert_called_once_with()
            workbook.Save.assert_called_once_with()
            workbook.Close.assert_called_once_with(SaveChanges=False)
            excel.Quit.assert_called_once_with()

    def test_round_nearest_whole_for_form9_grade11_grades(self):
        self.assertEqual(round_nearest_whole(84.49), 84)
        self.assertEqual(round_nearest_whole(84.5), 85)
        self.assertEqual(round_nearest_whole("89.75"), 90)
        self.assertEqual(round_nearest_whole("NOT A GRADE"), "NOT A GRADE")

    def test_lrns_preserve_leading_zeroes_and_ambiguous_names(self):
        self.assertEqual(norm_lrn(" 00123 "), "00123")
        self.assertEqual(norm_lrn("123.0"), "123")
        self.assertEqual(
            unique_name_lrns([
                {"name": "Same Name", "lrn": "1001"},
                {"name": "Same Name", "lrn": "1002"},
                {"name": "Other Name", "lrn": "1003"},
            ]),
            {"OTHER NAME": "1003"},
        )

    def test_scans_every_workbook_and_matches_grades_by_lrn(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            g11_folder = os.path.join(temp_dir, "grade11")
            g12_folder = os.path.join(temp_dir, "grade12")
            output_folder = os.path.join(temp_dir, "TRACKNAME")
            os.makedirs(g11_folder)
            os.makedirs(g12_folder)
            template = os.path.join(temp_dir, "template.xlsx")
            Workbook().save(template)

            students = [
                {"lrn": "1001", "source_name": "Shared Name",
                 "output_name": "Shared Name", "sex": "FEMALE", "grade": 81},
                {"lrn": "1002", "source_name": "Shared Name",
                 "output_name": "Shared Name", "sex": "MALE", "grade": 92},
            ]
            for index, student in enumerate(students, 1):
                _make_workbook(
                    os.path.join(g11_folder, f"grade11_{index}.xlsx"),
                    [student], student["grade"], True,
                )
                _make_workbook(
                    os.path.join(g12_folder, f"grade12_{index}.xlsx"),
                    [student], student["grade"] + 10, False,
                )

            with patch("services.sf10_engine.TemplatePatcher") as patcher:
                patcher.return_value.build_bytes.side_effect = (
                    lambda front, back: repr((front, back)).encode("utf-8")
                )
                engine = SF10Engine(g11_folder, g12_folder, template, lambda *_: None)
                engine.run()
                count = engine.save_to(output_folder)

            self.assertEqual(count, 2)
            self.assertEqual(engine.stats["success"], 2)
            self.assertTrue(os.path.isdir(os.path.join(output_folder, "MALE")))
            self.assertTrue(os.path.isdir(os.path.join(output_folder, "FEMALE")))
            female_files = os.listdir(os.path.join(output_folder, "FEMALE"))
            male_files = os.listdir(os.path.join(output_folder, "MALE"))
            self.assertEqual(len(female_files), 1)
            self.assertEqual(len(male_files), 1)
            with open(os.path.join(output_folder, "FEMALE", female_files[0]), "rb") as f:
                self.assertIn(b"'AT31': 81", f.read())
            with open(os.path.join(output_folder, "MALE", male_files[0]), "rb") as f:
                self.assertIn(b"'AT31': 92", f.read())

    def test_form9_routes_outputs_by_extracted_sex(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            g11_folder = os.path.join(temp_dir, "grade11")
            g12_folder = os.path.join(temp_dir, "grade12")
            output_folder = os.path.join(temp_dir, "TRACKNAME")
            os.makedirs(g11_folder)
            os.makedirs(g12_folder)
            template = os.path.join(temp_dir, "template.xlsx")
            Workbook().save(template)
            students = [
                {"lrn": "2001", "source_name": "Shared Name", "sex": "FEMALE", "grade": 71,
                 "last": "doe", "first": "jane", "middle": "marie"},
                {"lrn": "2002", "source_name": "Shared Name", "sex": "MALE", "grade": 82,
                 "last": "smith", "first": "john", "middle": None},
            ]
            for index, student in enumerate(students, 1):
                _make_form9_workbook(
                    os.path.join(g11_folder, f"grade11_{index}.xlsx"),
                    student, student["grade"], True,
                )
                _make_form9_workbook(
                    os.path.join(g12_folder, f"grade12_{index}.xlsx"),
                    student, student["grade"] + 10, False,
                )

            with (
                patch("services.form9_engine.SingleSheetPatcher") as patcher,
                patch("services.form9_engine._recalculate_workbooks") as recalculate,
            ):
                patcher.return_value.build_bytes.side_effect = (
                    lambda writes: repr(writes).encode("utf-8")
                )
                engine = Form9Engine(g11_folder, g12_folder, template, lambda *_: None)
                engine.run()
                count = engine.save_to(output_folder)

            self.assertEqual(count, 2)
            self.assertEqual(engine.stats["success"], 2)
            recalculated_paths = recalculate.call_args.args[0]
            self.assertEqual(len(recalculated_paths), 2)
            self.assertTrue(all(os.path.isfile(path) for path in recalculated_paths))
            female_files = os.listdir(os.path.join(output_folder, "FEMALE"))
            male_files = os.listdir(os.path.join(output_folder, "MALE"))
            self.assertEqual(len(female_files), 1)
            self.assertEqual(len(male_files), 1)
            with open(os.path.join(output_folder, "FEMALE", female_files[0]), "rb") as f:
                female_output = f.read()
                self.assertIn(b"'D14': 71", female_output)
                self.assertIn(b"'C6': 'DOE, JANE MARIE'", female_output)
            with open(os.path.join(output_folder, "MALE", male_files[0]), "rb") as f:
                male_output = f.read()
                self.assertIn(b"'D14': 82", male_output)
                self.assertIn(b"'C6': 'SMITH, JOHN'", male_output)


if __name__ == "__main__":
    unittest.main()
