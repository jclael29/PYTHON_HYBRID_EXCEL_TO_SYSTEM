"""
Lahat ng settings at cell maps nandito. Edit dito kung may nagbago sa template o sa Excel files.
"""
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

APP_TITLE = "Trace College - Form Automation"
TEMPLATE_NAME = "GAS_SHS_TEMPLATE_SF10.xlsx"
TEMPLATE_GLOB = "*SHS*TEMPLATE*SF10*.xlsx"   # panghanap kung bahagyang iba ang pangalan
OUTPUT_FOLDER_NAME = "TRACKNAME"      # pangalan ng main output folder; pwede itong palitan sa Save As

# MASTER SHEET (row 8 pababa) - student info galing sa GRADE 12 conso (hindi na sa Grade 11)
MASTER_START_ROW = 8
MASTER_COLS = {
    "name": "B",
    "lrn": "C",
    "sex": "D",
    "age": "E",
    "birthday": "F",
    "section": "G",       # section ng Grade 12 -> BACK (2)
    "last": "H",
    "first": "I",
    "middle": "J",
    "school_year": "N",
    "gwa": "K",
    "last_school": "L",
    "school_addres": "M",   
}

# Grade 11 MASTER SHEET: LRN (pang-match) at SECTION lang ang kinukuha
G11_MASTER_COLS = {"name": "B", "lrn": "C", "section": "G"}

# Mga field na dapat TEXT (hindi ginagawang numero)
TEXT_KEYS = {"name", "sex", "last", "first", "middle", "section", "last_school",
             "school_addres", "school_year"}
NAME_KEYS = ("last", "first", "middle")

# MASTER SHEET -> FRONT
FRONT_INFO = {
    "last": "F8",
    "first": "Y8",
    "middle": "AZ8",
    "lrn": "C9",
    "birthday": "AA9",
    "sex": "AN9",
    "school_year": "BA23",
    "gwa": "AH13",
    "last_school": "Z14",
    "school_addres": "AW14",
}

# SECTION: dalawang cell bawat isa
FRONT_SECTION_CELLS = ("AS25", "AS68")   # section galing Grade 11 -> FRONT
BACK_SECTION_CELLS = ("AS5", "AS28" , "AS47")     # section galing Grade 12 -> BACK (2)

# G11 grade sheets (pangalan nasa B6 pababa, grades C/D pairs)
G11_START_ROW = 6
G11_PAIRS = [("C", "D"), ("F", "G"), ("I", "J"), ("L", "M"), ("O", "P"),
             ("R", "S"), ("U", "V"), ("X", "Y"), ("AA", "AB"), ("AD", "AE")]
G11_SHEETS = {
    "sem1": {"name": "1ST SEM CONSO", "index": 7, "pairs": G11_PAIRS,      "first_row": 31},
    "sem2": {"name": "2ND SEM CONSO", "index": 8, "pairs": G11_PAIRS[:9],  "first_row": 74},
}
FRONT_GRADE_COLS = ("AT", "AY")

# G12 grade sheets (B8 pababa) -> BACK (2)
G12_START_ROW = 8
G12_SHEETS = {
    "1ST T": {"map": [("C", "AY11"), ("D", "AY12"), ("E", "AY13"),
                      ("F", "AY14"), ("G", "AY15"), ("H", "AY16")]},
    "2ND T": {"map": [("C", "AY34"), ("D", "AY35"), ("E", "AY36"),
                      ("F", "AY37"), ("G", "AY38")]},
    "3RD T": {"map": [("C", "AY54"), ("D", "AY56"), ("E", "AY57"),
                      ("F", "AY58")]},
}