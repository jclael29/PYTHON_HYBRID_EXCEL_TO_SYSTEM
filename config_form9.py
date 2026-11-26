"""
GRAD FORM 9 - lahat ng cell maps nandito. Edit dito kung may nagbago sa template o sa Excel files.
"""

TEMPLATE_NAME = "FORM_9_TEMPLATE.xlsx"
TEMPLATE_KEYWORDS = ("FORM", "9", "TEMPLATE")   # panghanap kung iba ang eksaktong pangalan ng file
OUTPUT_FOLDER_NAME = "TRACKNAME"                # pangalan ng main output folder
PAGE_TITLE = "Form 9"

# ---- Student info: GRADE 12 conso, sheet MASTER SHEET (pinaka unang sheet), row 8 pababa ----
MASTER_START_ROW = 8
MASTER_COLS = {
    "last": "H",
    "first": "I",
    "middle": "J",
    "lrn": "C",
    "sex": "D",
    "birthday": "F",
    "last_school": "M",
    "school_year": "N",
    "father": "P",
    "mother": "Q",
}

# Grade 11 MASTER SHEET: LRN (pang-match) lang ang kailangan
G11_MASTER_COLS = {"name": "B", "lrn": "C"}

# Mga field na dapat TEXT (hindi ginagawang numero)
TEXT_KEYS = {
    "name", "last", "first", "middle", "sex", "father", "mother",
    "last_school", "school_year",
}

# MASTER SHEET -> FORM 9 TEMPLATE (isa lang ang sheet)
FORM9_INFO = {
    "name": "C6",
    "father": "C7",
    "last_school": "C8",
    "lrn": "J5",
    "birthday": "J6",
    "mother": "J7",
    "school_year": "J8",
}

# ---- GRADE 11 conso (pangalan nasa column B, row 6 pababa) ----
G11_SHEETS = {
    "sem1": {
        "name": "1ST SEM CONSO", "index": 7, "start_row": 6,       # pang-8 na sheet
        "map": [("E", "D14"), ("H", "D15"), ("K", "D16"), ("N", "D17"), ("Q", "D18"),
                ("T", "D20"), ("W", "D22"), ("Z", "D23"), ("AC", "D25"), ("AF", "D26")],
    },
    "sem2": {
        "name": "2ND SEM CONSO", "index": 8, "start_row": 6,       # pang-9 na sheet
        "map": [("E", "K14"), ("H", "K15"), ("K", "K16"), ("N", "K17"), ("Q", "K18"),
                ("T", "K19"), ("W", "K21"), ("Z", "K22"), ("AC", "K24")],
    },
}

# ---- GRADE 12 conso (pangalan nasa column B, row 8 pababa) ----
G12_SHEETS = {
    "1ST T": {"index": 3, "start_row": 8,                          # pang-4 na sheet
              "map": [("C", "D34"), ("D", "D36"), ("E", "D38"),
                      ("F", "D39"), ("G", "D41"), ("H", "D42")]},
    "2ND T": {"index": 4, "start_row": 8,                          # pang-5 na sheet
              "map": [("C", "K34"), ("D", "K35"), ("E", "K36"),
                      ("F", "K38"), ("G", "K40")]},
    "3RD T": {"index": 5, "start_row": 8,                          # pang-6 na sheet
              "map": [("C", "D51"), ("D", "D53"), ("E", "D55"),
                      ("F", "D57")]},
}