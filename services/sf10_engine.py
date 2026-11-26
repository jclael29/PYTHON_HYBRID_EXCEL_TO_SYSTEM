"""
SF10 ENGINE (walang tkinter - pwedeng i-test mag-isa)
Grade 11 at Grade 12 folders -> SF10 template, isang file kada student.
"""
import os

from openpyxl import load_workbook

import config
from services.excel_utils import (
    check_xlsx, clean, clean_text, col_get, find_sheet, fmt_birthday,
    list_excel_files, norm_lrn, norm_name, safe_filename, sex_output_folder,
    unique_name_lrns,
)
from services.template_patcher import TemplatePatcher


class SF10Engine:
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
    def _read_master(self, wb, cols=None):
        cols = cols or config.MASTER_COLS
        ws = wb.worksheets[0]  # MASTER SHEET = pinaka unang sheet
        students = []
        for row in ws.iter_rows(min_row=config.MASTER_START_ROW, values_only=True):
            name = clean(col_get(row, cols["name"]))
            lrn = clean(col_get(row, cols["lrn"]))
            if name is None and lrn is None:
                continue
            rec = {k: (clean_text if k in config.TEXT_KEYS or k == "lrn" else clean)(col_get(row, c))
                   for k, c in cols.items()}
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
        check_xlsx(self.template_path, "SF10 template")
        self.patcher = TemplatePatcher(self.template_path)

        g11_by_lrn, g11_grades_by_lrn = {}, {}
        for index, path in enumerate(g11_files, 1):
            self._say(2, f"Binabasa ang Grade 11 file {index}/{len(g11_files)}...")
            wb = load_workbook(path, data_only=True, read_only=True)
            try:
                students = self._read_master(wb, config.G11_MASTER_COLS)
                for student in students:
                    if student["lrn"]:
                        existing = g11_by_lrn.setdefault(student["lrn"], student)
                        if existing is not student:
                            for field, value in student.items():
                                if existing.get(field) is None and value is not None:
                                    existing[field] = value
                name_to_lrn = unique_name_lrns(students)
                for key, sheet_config in config.G11_SHEETS.items():
                    ws = find_sheet(wb, sheet_config["name"], sheet_config["index"])
                    sheet_grades = self._read_grade_sheet(ws, config.G11_START_ROW, name_to_lrn)
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
                students = self._read_master(wb)
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
                for key in config.G12_SHEETS:
                    ws = find_sheet(wb, key, None)
                    sheet_grades = self._read_grade_sheet(ws, config.G12_START_ROW, name_to_lrn)
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
                self.log.append(f"SKIPPED  {s11['name']}: nasa Grade 11 pero wala sa Grade 12 (walang SF10)")

        self._say(100, "Tapos na. Pindutin ang Save para i-save ang mga file.")

    def save_to(self, folder):
        """Isulat sa disk ang lahat ng na-process (tinatawag lang kapag pinindot ang Save)."""
        os.makedirs(folder, exist_ok=True)
        for subfolder in ("MALE", "FEMALE"):
            os.makedirs(os.path.join(folder, subfolder), exist_ok=True)
        for subfolder, name, data in self.results:
            with open(os.path.join(folder, subfolder, name), "wb") as f:
                f.write(data)
        with open(os.path.join(folder, "processing_log.txt"), "w", encoding="utf-8") as f:
            f.write("\n".join(self.log) if self.log else "Walang problema.")
        return len(self.results)

    def _make_one(self, st, g11_by_lrn, g11_grades_by_lrn, g12_grades_by_lrn, used_names):
        missing = []
        front, back = {}, {}

        # dapat TEXT ang apelyido/pangalan/middle - kung numero, mali ang column
        for key in config.NAME_KEYS:
            v = st.get(key)
            if v is not None and not isinstance(v, str):
                raise ValueError(
                    f'ang "{key}" (column {config.MASTER_COLS[key]} sa Grade 12 MASTER SHEET) ay numero: {v}. '
                    f'Mali yata ang column sa MASTER_COLS.')

        # student info (galing sa Grade 12 MASTER SHEET) -> FRONT
        for key, coord in config.FRONT_INFO.items():
            v = st.get(key)
            if key == "birthday":
                v = fmt_birthday(v)
            if v is not None:
                front[coord] = v

        # section ng Grade 12 -> BACK (2)
        if st.get("section") is not None:
            for coord in config.BACK_SECTION_CELLS:
                back[coord] = st["section"]

        # G11 grades -> FRONT. Hanapin ang student sa G11 gamit ang LRN,
        # tapos ang pangalan niya sa G11 ang gagamitin sa 1ST/2ND SEM CONSO.
        g11 = g11_by_lrn.get(st["lrn"]) if st["lrn"] else None
        if g11 is None:
            missing.append("walang Grade 11 (LRN)")
        elif g11.get("section") is not None:
            for coord in config.FRONT_SECTION_CELLS:   # section ng Grade 11 -> FRONT
                front[coord] = g11["section"]
        g11_grades = g11_grades_by_lrn.get(st["lrn"], {})
        for sem, cfg in config.G11_SHEETS.items():
            row = g11_grades.get(sem)
            if row is None:
                if g11 is not None:
                    missing.append(f"walang {cfg['name']}")
                continue
            for i, (a, b) in enumerate(cfg["pairs"]):
                r = cfg["first_row"] + i
                for src, dst in zip((a, b), config.FRONT_GRADE_COLS):
                    v = clean(col_get(row, src))
                    if v is not None:
                        front[f"{dst}{r}"] = v

        # G12 grades -> BACK (pangalan ng student sa G12 file mismo)
        for sheet, cfg in config.G12_SHEETS.items():
            row = g12_grades_by_lrn.get(st["lrn"], {}).get(sheet)
            if row is None:
                missing.append(f"walang {sheet}")
                continue
            for src, dst in cfg["map"]:
                v = clean(col_get(row, src))
                if v is not None:
                    back[dst] = v

        # i-build sa memory (hindi pa isinusulat sa disk)
        output_group = sex_output_folder(st.get("sex"), st.get("name") or st["lrn"])

        base = ", ".join(p for p in [st.get("last")] if p)
        rest = " ".join(p for p in [st.get("first"), st.get("middle")] if p)
        fname = safe_filename(f"{base}, {rest}" if base and rest else (st["name"] or st["lrn"]))
        final, n = fname, 2
        while final.upper() in used_names:
            final = f"{fname} ({n})"
            n += 1
        used_names.add(final.upper())
        self.results.append((output_group, final + ".xlsx", self.patcher.build_bytes(front, back)))
        return missing