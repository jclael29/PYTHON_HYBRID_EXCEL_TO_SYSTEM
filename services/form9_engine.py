"""
FORM 9 ENGINE (walang tkinter - pwedeng i-test mag-isa)
Grade 11 conso + Grade 12 conso -> FORM 9 TEMPLATE (isang sheet), isang file kada student.
Pareho ang daloy ng SF10; iba lang ang mga target na cell (nasa config_form9.py).
"""
import os

from openpyxl import load_workbook

import config_form9 as cfg
from services.excel_utils import (
    check_xlsx, clean, clean_text, col_get, find_sheet, fmt_birthday,
    norm_lrn, norm_name, safe_filename,
)
from services.template_patcher import SingleSheetPatcher


class Form9Engine:
    def __init__(self, g11_path, g12_path, template_path, report):
        self.g11_path, self.g12_path = g11_path, g12_path
        self.template_path = template_path
        self.report = report  # report(percent, status_text, stats_dict)
        self.stats = {"processed": 0, "success": 0, "missing": 0, "failed": 0}
        self.log = []
        self.results = []     # [(filename, bytes)] - nasa memory pa, hindi pa naka-save

    def _say(self, pct, text):
        self.report(pct, text, dict(self.stats))

    # ---- readers ----
    def _read_master(self, wb, cols):
        ws = wb.worksheets[0]  # MASTER SHEET = pinaka unang sheet
        students = []
        for row in ws.iter_rows(min_row=cfg.MASTER_START_ROW, values_only=True):
            name = clean(col_get(row, cols["name"]))
            lrn = clean(col_get(row, cols["lrn"]))
            if name is None and lrn is None:
                continue
            rec = {k: (clean_text if k in cfg.TEXT_KEYS else clean)(col_get(row, c))
                   for k, c in cols.items()}
            rec["lrn"] = norm_lrn(rec["lrn"])
            students.append(rec)
        return students

    def _read_grade_sheet(self, ws, start_row):
        """name(normalized) -> row tuple"""
        out = {}
        for row in ws.iter_rows(min_row=start_row, values_only=True):
            key = norm_name(col_get(row, "B"))
            if key and key not in out:
                out[key] = row
        return out

    # ---- main ----
    def run(self):
        check_xlsx(self.g11_path, "Grade 11 file")
        check_xlsx(self.g12_path, "Grade 12 file")
        check_xlsx(self.template_path, "Form 9 template")
        self.patcher = SingleSheetPatcher(self.template_path)

        self._say(2, "Binabasa ang Grade 11 file...")
        wb11 = load_workbook(self.g11_path, data_only=True, read_only=True)
        students11 = self._read_master(wb11, cfg.G11_MASTER_COLS)   # LRN lang ang kailangan
        g11_by_lrn = {s["lrn"]: s for s in students11 if s["lrn"]}
        sems = {}
        for key, sc in cfg.G11_SHEETS.items():
            ws = find_sheet(wb11, sc["name"], sc["index"])
            sems[key] = self._read_grade_sheet(ws, sc["start_row"])
        wb11.close()

        self._say(6, "Binabasa ang Grade 12 file...")
        wb12 = load_workbook(self.g12_path, data_only=True, read_only=True)
        students12 = self._read_master(wb12, cfg.MASTER_COLS)       # dito galing ang student info
        g12_sheets = {}
        for key, sc in cfg.G12_SHEETS.items():
            ws = find_sheet(wb12, key, sc["index"])
            g12_sheets[key] = self._read_grade_sheet(ws, sc["start_row"])
        wb12.close()

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
                missing = self._make_one(st, sems, g11_by_lrn, g12_sheets, used_names)
                self.stats["success"] += 1
                if missing:
                    self.stats["missing"] += 1
                    self.log.append(f"MISSING  {label}: {', '.join(missing)}")
            except Exception as e:  # noqa
                self.stats["failed"] += 1
                self.log.append(f"FAILED   {label}: {e}")

        # nasa Grade 11 pero wala sa Grade 12 -> hindi na-generate, ilista lang
        g12_lrns = {s["lrn"] for s in students12 if s["lrn"]}
        for s11 in students11:
            if s11["lrn"] and s11["lrn"] not in g12_lrns:
                self.log.append(f"SKIPPED  {s11['name']}: nasa Grade 11 pero wala sa Grade 12 (walang Form 9)")

        self._say(100, "Tapos na. Pindutin ang Save para i-save ang mga file.")

    def save_to(self, folder):
        """Isulat sa disk ang lahat ng na-process (tinatawag lang kapag pinindot ang Save)."""
        os.makedirs(folder, exist_ok=True)
        for name, data in self.results:
            with open(os.path.join(folder, name), "wb") as f:
                f.write(data)
        with open(os.path.join(folder, "processing_log.txt"), "w", encoding="utf-8") as f:
            f.write("\n".join(self.log) if self.log else "Walang problema.")
        return len(self.results)

    def _make_one(self, st, sems, g11_by_lrn, g12_sheets, used_names):
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
        keys = []
        if g11 is not None:
            keys.append(norm_name(g11["name"]))
        keys.append(norm_name(st["name"]))  # panghuli: pangalan ng G12
        for sem, sc in cfg.G11_SHEETS.items():
            row = next((sems[sem][k] for k in keys if k in sems[sem]), None)
            if row is None:
                if g11 is not None:
                    missing.append(f"walang {sc['name']}")
                continue
            for src, dst in sc["map"]:
                v = clean(col_get(row, src))
                if v is not None:
                    writes[dst] = v

        # G12 grades (pangalan ng student sa G12 file mismo)
        gkey = norm_name(st["name"])
        for sheet, sc in cfg.G12_SHEETS.items():
            row = g12_sheets[sheet].get(gkey)
            if row is None:
                missing.append(f"walang {sheet}")
                continue
            for src, dst in sc["map"]:
                v = clean(col_get(row, src))
                if v is not None:
                    writes[dst] = v

        # i-build sa memory (hindi pa isinusulat sa disk)
        fname = safe_filename(str(st["name"] or st["lrn"]))
        final, n = fname, 2
        while final.upper() in used_names:
            final = f"{fname} ({n})"
            n += 1
        used_names.add(final.upper())
        self.results.append((final + ".xlsx", self.patcher.build_bytes(writes)))
        return missing