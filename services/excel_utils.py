"""Maliliit na helper para sa pagbasa/paglinis ng Excel data."""
import os
import re
import zipfile
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
    return re.sub(r"\s+", "", str(v)).replace(".0", "") if str(v).strip() else ""


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
