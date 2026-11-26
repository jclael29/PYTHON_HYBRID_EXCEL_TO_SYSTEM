"""Maliliit na helper para sa pagbasa/paglinis ng Excel data."""
import os
import re
import zipfile
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from datetime import datetime, date

from openpyxl.utils import column_index_from_string


def clean(v):
    """Linisin ang value; ginagawang numero ang numeric-looking na text."""
    if v is None:
        return None
    if isinstance(v, str):
        v = v.strip()
        if v == "" or v.startswith("#"):
            return None
        try:
            f = float(v)
            return int(f) if f.is_integer() else f
        except ValueError:
            return v
    return v


def round_nearest_whole(v):
    """Round a numeric grade to the nearest whole number (0.5 rounds up)."""
    v = clean(v)
    if v is None or isinstance(v, bool):
        return v
    try:
        return int(Decimal(str(v)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    except (InvalidOperation, ValueError):
        return v


def clean_text(v):
    """Para sa text na field: hindi ginagawang numero ang text."""
    if v is None:
        return None
    if isinstance(v, str):
        v = v.strip()
        return None if v == "" or v.startswith("#") else v
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return v


def norm_name(s):
    if s is None:
        return ""
    s = re.sub(r"[^\w ]", " ", str(s).upper().replace("_", " "))
    return re.sub(r"\s+", " ", s).strip()


def norm_lrn(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    value = re.sub(r"\s+", "", str(v).strip())
    return value[:-2] if value.endswith(".0") else value


def sex_output_folder(value, student_label):
    sex = str(value or "").strip().upper()
    if sex in ("M", "MALE"):
        return "MALE"
    if sex in ("F", "FEMALE"):
        return "FEMALE"
    raise ValueError(f'Hindi matukoy ang sex ni {student_label}: {sex or "walang value"}')


def col_get(row, col):
    i = column_index_from_string(col) - 1
    return row[i] if i < len(row) else None


def find_sheet(wb, name, index):
    for ws in wb.worksheets:
        if ws.title.strip().upper() == name.strip().upper():
            return ws
    if index is not None and index < len(wb.worksheets):
        return wb.worksheets[index]
    raise KeyError(f'Walang sheet na "{name}"')


def safe_filename(s):
    return re.sub(r'[\\/:*?"<>|]', "", s).strip() or "SF10"


def fmt_birthday(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%m/%d/%Y")
    return v


def check_xlsx(path, label):
    """Linawin kung bakit hindi mabuksan ang file (imbes na teknikal na error)."""
    try:
        with open(path, "rb") as f:
            head = f.read(8)
    except OSError as e:
        raise ValueError(f"Hindi mabuksan ang {label}:\n{os.path.basename(path)}\n{e}")
    name = os.path.basename(path)
    if head[:2] == b"PK":
        try:
            with zipfile.ZipFile(path) as z:
                names = [n.lower() for n in z.namelist()]
        except zipfile.BadZipFile:
            raise ValueError(f"Sira o hindi kumpleto ang {label}:\n{name}\n\nBuksan sa Excel, tapos Save As .xlsx.")
        if "[content_types].xml" in names:
            return
        if any(n.endswith((".xlsx", ".xlsm")) for n in names):
            raise ValueError(f"Ang {label} ay zip archive na may Excel file sa loob:\n{name}\n\nI-extract muna ito, tapos yung Excel file ang piliin.")
        raise ValueError(f"Hindi totoong .xlsx ang {label}:\n{name}\n\nBuksan sa Excel, tapos Save As .xlsx (Excel Workbook), at yun ang piliin.")
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        raise ValueError(f"Ang {label} ay luma (.xls) na format o may password:\n{name}\n\nBuksan sa Excel (alisin ang password kung meron), tapos Save As .xlsx.")
    raise ValueError(f"Hindi Excel file ang {label}:\n{name}")


def list_excel_files(folder, label):
    """Ibalik ang lahat ng Excel workbook sa napiling folder."""
    if not os.path.isdir(folder):
        raise ValueError(f"Hindi makita ang {label} folder:\n{folder}")
    try:
        paths = sorted(
            (
                os.path.join(folder, name)
                for name in os.listdir(folder)
                if name.lower().endswith((".xlsx", ".xlsm"))
                and not name.startswith("~$")
                and os.path.isfile(os.path.join(folder, name))
            ),
            key=str.lower,
        )
    except OSError as e:
        raise ValueError(f"Hindi mabasa ang {label} folder:\n{folder}\n{e}")
    if not paths:
        raise ValueError(f"Walang Excel file sa {label} folder:\n{folder}")
    return paths


def unique_name_lrns(students):
    """I-map ang natatanging pangalan sa LRN; alisin ang mga pangalang hindi ligtas i-match."""
    matches = {}
    ambiguous = set()
    for student in students:
        name, lrn = norm_name(student.get("name")), student.get("lrn")
        if not name or not lrn:
            continue
        if name in matches and matches[name] != lrn:
            ambiguous.add(name)
        else:
            matches[name] = lrn
    for name in ambiguous:
        matches.pop(name, None)
    return matches


def find_template(dirs, name, keywords):
    """Hanapin ang template: eksaktong pangalan muna, tapos kahit bahagyang iba ang pangalan
    (hal. may "Copy of") basta nasa pangalan ang lahat ng keywords. Hindi sensitive sa laki ng letra."""
    for d in dirs:
        if not d or not os.path.isdir(d):
            continue
        p = os.path.join(d, name)
        if os.path.exists(p):
            return p
        for f in sorted(os.listdir(d)):
            up = f.upper()
            if f.startswith("~$") or not up.endswith((".XLSX", ".XLSM")):
                continue
            if all(k in up for k in keywords):
                return os.path.join(d, f)
    return ""
