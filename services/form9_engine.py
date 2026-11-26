"""
FORM 9 ENGINE (walang tkinter - pwedeng i-test mag-isa)
Grade 11 at Grade 12 folders -> FORM 9 TEMPLATE (isang sheet), isang file kada student.
Pareho ang daloy ng SF10; iba lang ang mga target na cell (nasa config_form9.py).
"""
import os

from openpyxl import load_workbook

import config_form9 as cfg
from services.excel_utils import (
    check_xlsx, clean, clean_text, col_get, find_sheet, fmt_birthday,
    list_excel_files, norm_lrn, norm_name, safe_filename, sex_output_folder,
    round_nearest_whole, unique_name_lrns,
)
from services.template_patcher import SingleSheetPatcher


def _recalculate_workbooks(paths):
    """Force Excel to fully recalculate and save each generated Form 9 workbook."""
    paths = tuple(paths)
    if not paths:
        return

    try:
        from win32com.client import DispatchEx
    except ImportError as exc:
        raise OSError("Form 9 recalculation requires pywin32 and Microsoft Excel.") from exc

    excel = None
    try:
        excel = DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        for path in paths:
            workbook = None
            try:
                workbook = excel.Workbooks.Open(
                    os.path.abspath(path), UpdateLinks=0, ReadOnly=False
                )
                excel.CalculateFullRebuild()
                workbook.Save()
            except Exception as exc:
                raise OSError(
                    f"Could not fully recalculate Form 9 workbook: {path}"
                ) from exc
            finally:
                if workbook is not None:
                    workbook.Close(SaveChanges=False)
    except OSError:
        raise
    except Exception as exc:
        raise OSError("Could not start Excel to recalculate Form 9 workbooks.") from exc
    finally:
        if excel is not None:
            excel.Quit()


class Form9Engine:
    def __init__(self, g11_path, g12_path, template_path, report):
        self.g11_path, self.g12_path = g11_path, g12_path
        self.template_path = template_path
        self.report = report  # report(percent, status_text, stats_dict)
        self.stats = {"processed": 0, "success": 0, "missing": 0, "failed": 0}
        self.log = []
        self.results = []     # [(sex_folder, filename, bytes)] - nasa memory pa

    def _say(self, pct, text):
        self.report(pct, text, dict(self.stats))

    # ---- readers ----
    def _read_master(self, wb, cols):
        ws = wb.worksheets[0]  # MASTER SHEET = pinaka unang sheet
        students = []
        for row in ws.iter_rows(min_row=cfg.MASTER_START_ROW, values_only=True):
            lrn = clean(col_get(row, cols["lrn"]))
            rec = {k: (clean_text if k in cfg.TEXT_KEYS or k == "lrn" else clean)(col_get(row, c))
                   for k, c in cols.items()}
            if "name" not in rec:
                parts = {
                    key: str(rec[key]).strip().upper() if rec.get(key) is not None else ""
                    for key in ("last", "first", "middle")
                }
                last = parts["last"]
                given = " ".join(part for part in (parts["first"], parts["middle"]) if part)
                rec["name"] = (
                    f"{last}, {given}" if last and given else last or given
                ) or None
            if rec["name"] is None and lrn is None:
                continue
            rec["lrn"] = norm_lrn(rec["lrn"])
            students.append(rec)
        return students

    def _read_grade_sheet(self, ws, start_row, name_to_lrn):
        """LRN -> row tuple, using the matching workbook's master sheet."""
        out = {}
        for row in ws.iter_rows(min_row=start_row, values_only=True):
            key = norm_name(col_get(row, "B"))
            lrn = name_to_lrn.get(key)
            if lrn and lrn not in out:
                out[lrn] = row
        return out

    # ---- main ----
    def run(self):
        g11_files = list_excel_files(self.g11_path, "Grade 11")
        g12_files = list_excel_files(self.g12_path, "Grade 12")
        for path in g11_files:
            check_xlsx(path, "Grade 11 file")
        for path in g12_files:
            check_xlsx(path, "Grade 12 file")
        check_xlsx(self.template_path, "Form 9 template")
        self.patcher = SingleSheetPatcher(self.template_path)

        g11_by_lrn, g11_grades_by_lrn = {}, {}
        for index, path in enumerate(g11_files, 1):
            self._say(2, f"Binabasa ang Grade 11 file {index}/{len(g11_files)}...")
            wb = load_workbook(path, data_only=True, read_only=True)
            try:
                students = self._read_master(wb, cfg.G11_MASTER_COLS)
                for student in students:
                    if student["lrn"]:
                        existing = g11_by_lrn.setdefault(student["lrn"], student)
                        if existing is not student:
                            for field, value in student.items():
                                if existing.get(field) is None and value is not None:
                                    existing[field] = value
                name_to_lrn = unique_name_lrns(students)
                for key, sc in cfg.G11_SHEETS.items():
                    ws = find_sheet(wb, sc["name"], sc["index"])
                    sheet_grades = self._read_grade_sheet(ws, sc["start_row"], name_to_lrn)
                    for lrn, row in sheet_grades.items():
                        g11_grades_by_lrn.setdefault(lrn, {}).setdefault(key, row)
            finally:
                wb.close()

        students12_by_lrn, students12_without_lrn = {}, []
        g12_grades_by_lrn = {}
        for index, path in enumerate(g12_files, 1):
            self._say(6, f"Binabasa ang Grade 12 file {index}/{len(g12_files)}...")
            wb = load_workbook(path, data_only=True, read_only=True)
            try:
                students = self._read_master(wb, cfg.MASTER_COLS)
                name_to_lrn = unique_name_lrns(students)
                for student in students:
                    lrn = student["lrn"]
                    if not lrn:
                        students12_without_lrn.append(student)
                        continue
                    existing = students12_by_lrn.setdefault(lrn, student)
                    if existing is not student:
                        for field, value in student.items():
                            if existing.get(field) is None and value is not None:
                                existing[field] = value
                for key, sc in cfg.G12_SHEETS.items():
                    ws = find_sheet(wb, key, sc["index"])
                    sheet_grades = self._read_grade_sheet(ws, sc["start_row"], name_to_lrn)
                    for lrn, row in sheet_grades.items():
                        g12_grades_by_lrn.setdefault(lrn, {}).setdefault(key, row)
            finally:
                wb.close()
        students12 = list(students12_by_lrn.values()) + students12_without_lrn

        total = len(students12)
        if total == 0:
            raise ValueError("Walang student na nabasa sa MASTER SHEET ng Grade 12 (row 8 pababa).")

        used_names = set()
        for i, st in enumerate(students12, 1):
            label = st["name"] or st["lrn"]
            pct = 10 + int(88 * (i - 1) / total)
            self._say(pct, f"Processing {i}/{total}: {label}")
            self.stats["processed"] += 1
            try:
                missing = self._make_one(st, g11_by_lrn, g11_grades_by_lrn,
                                         g12_grades_by_lrn, used_names)
                self.stats["success"] += 1
                if missing:
                    self.stats["missing"] += 1
                    self.log.append(f"MISSING  {label}: {', '.join(missing)}")
            except Exception as e:  # noqa
                self.stats["failed"] += 1
                self.log.append(f"FAILED   {label}: {e}")

        # nasa Grade 11 pero wala sa Grade 12 -> hindi na-generate, ilista lang
        g12_lrns = {s["lrn"] for s in students12 if s["lrn"]}
        for s11 in g11_by_lrn.values():
            if s11["lrn"] and s11["lrn"] not in g12_lrns:
                self.log.append(f"SKIPPED  {s11['name']}: nasa Grade 11 pero wala sa Grade 12 (walang Form 9)")

        self._say(100, "Tapos na. Pindutin ang Save As para i-save ang mga file.")

    def save_to(self, folder):
        """Isulat sa disk ang lahat ng na-process (tinatawag lang kapag pinindot ang Save As)."""
        os.makedirs(folder, exist_ok=True)
        for subfolder in ("MALE", "FEMALE"):
            os.makedirs(os.path.join(folder, subfolder), exist_ok=True)
        for subfolder, name, data in self.results:
            with open(os.path.join(folder, subfolder, name), "wb") as f:
                f.write(data)
        _recalculate_workbooks([
            os.path.join(folder, subfolder, name)
            for subfolder, name, _ in self.results
        ])
        with open(os.path.join(folder, "processing_log.txt"), "w", encoding="utf-8") as f:
            f.write("\n".join(self.log) if self.log else "Walang problema.")
        return len(self.results)

    def _make_one(self, st, g11_by_lrn, g11_grades_by_lrn, g12_grades_by_lrn, used_names):
        missing = []
        writes = {}

        # student info (galing sa Grade 12 MASTER SHEET) -> FORM 9
        for key, coord in cfg.FORM9_INFO.items():
            v = st.get(key)
            if key == "birthday":
                v = fmt_birthday(v)
            if v is not None:
                writes[coord] = v

        # G11 grades. Hanapin ang student sa G11 gamit ang LRN,
        # tapos ang pangalan niya sa G11 ang gagamitin sa 1ST/2ND SEM CONSO.
        g11 = g11_by_lrn.get(st["lrn"]) if st["lrn"] else None
        if g11 is None:
            missing.append("walang Grade 11 (LRN)")
        g11_grades = g11_grades_by_lrn.get(st["lrn"], {})
        for sem, sc in cfg.G11_SHEETS.items():
            row = g11_grades.get(sem)
            if row is None:
                if g11 is not None:
                    missing.append(f"walang {sc['name']}")
                continue
            for src, dst in sc["map"]:
                v = round_nearest_whole(col_get(row, src))
                if v is not None:
                    writes[dst] = v

        # G12 grades (pangalan ng student sa G12 file mismo)
        for sheet, sc in cfg.G12_SHEETS.items():
            row = g12_grades_by_lrn.get(st["lrn"], {}).get(sheet)
            if row is None:
                missing.append(f"walang {sheet}")
                continue
            for src, dst in sc["map"]:
                v = clean(col_get(row, src))
                if v is not None:
                    writes[dst] = v

        output_group = sex_output_folder(st.get("sex"), st.get("name") or st["lrn"])

        # i-build sa memory (hindi pa isinusulat sa disk)
        fname = safe_filename(str(st["name"] or st["lrn"]))
        final, n = fname, 2
        while final.upper() in used_names:
            final = f"{fname} ({n})"
            n += 1
        used_names.add(final.upper())
        self.results.append((output_group, final + ".xlsx", self.patcher.build_bytes(writes)))
        return missing