import os
import re
import json
from pathlib import Path
from datetime import datetime
import openpyxl
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

# ============================================================
# EXCEL SKILL 2 - AUTO CHECKER
# DATA VALIDATION & NAME MANAGER
#
# 4 Modules | 10 Marks | 100+ Students
#
# Checks ACTUAL Excel objects:
#   1. Workbook Defined Names
#   2. List Data Validation
#   3. Dependent Data Validation
#   4. Whole Number / Date / Text Length validation
#
# Student's entered value is NOT used to decide correctness.
# The validation rule itself is checked.
#
# Possible outcomes are recorded as 0/1.
# ============================================================

BASE_DIR = Path(__file__).parent
SUBMISSIONS_DIR = BASE_DIR / "submissions"
RESULTS_DIR = BASE_DIR / "results"

# Module marks
NAMED_MANAGER_MARKS = 2.5
BASIC_DROPDOWN_MARKS = 2.5
ADVANCED_DROPDOWN_MARKS = 2.5
DATA_VALIDATION_MARKS = 2.5


# ============================================================
# GENERAL HELPERS
# ============================================================

def norm(value):
    if value is None:
        return ""

    return (
        re.sub(r"\s+", "", str(value))
        .upper()
        .replace("$", "")
        .lstrip("=")
        .replace("{", "")
        .replace("}", "")
    )


def dv_applies_to_cell(dv, cell):

    sqref = str(dv.sqref)

    for ref in sqref.split():

        # Exact cell
        if ref.upper() == cell.upper():
            return True

        # Range
        if ":" in ref:
            try:
                min_col, min_row, max_col, max_row = \
                    range_boundaries(ref)

                row, col = coordinate_to_tuple(cell)

                if (
                    min_row <= row <= max_row
                    and min_col <= col <= max_col
                ):
                    return True

            except Exception:
                pass

    return False


def get_dv(ws, cell):

    for dv in ws.data_validations.dataValidation:

        if dv_applies_to_cell(dv, cell):
            return dv

    return None


def print_all_dv(ws):

    dvs = ws.data_validations.dataValidation
    if not dvs:
        return

    for i, dv in enumerate(dvs, 1):
        pass


def check_c10(ws):

    dv = get_dv(ws, "C10")

    if dv is None:
        return (
            0,
            "No Data Validation found on C10"
        )

    if dv.type != "list":
        return (
            0,
            "C10 must use List validation"
        )

    actual = norm(dv.formula1)

    accepted = {
        norm("=$G$3:$H$3"),
        norm("$G$3:$H$3"),
        norm("=G3:H3"),
        norm("G3:H3"),
        norm("=Category"),
        norm("=CategoryList"),
        norm('"Electronics,Furniture"'),
        norm("Electronics,Furniture")
    }

    if actual in accepted:
        return 1, "Correct"

    return (
        0,
        f"Incorrect source on C10: {dv.formula1}"
    )


def check_c11(ws):
    value = ws["C11"].value

    if value is None:
        return 0, "C11 is empty"

    # Array Formula
    if hasattr(value, "text"):
        formula = value.text
    else:
        formula = str(value)

    formula = str(formula).strip()
    formula = formula.replace("{", "")
    formula = formula.replace("}", "")
    formula = formula.replace("$", "")
    formula = formula.replace(" ", "")
    formula = formula.upper()

    if formula in ["=INDIRECT(C10)", "INDIRECT(C10)"]:
        return 1, "Correct"

    return 0, f"Incorrect formula on C11: {formula}"


def check_list_dv(ws, cell, accepted_sources):
    if cell.upper() == "C10":
        return check_c10(ws)
    elif cell.upper() == "C11":
        return check_c11(ws)
    else:
        dv = get_dv(ws, cell)
        if dv is None:
            return 0, f"No Data Validation found on {cell}"
        if dv.type != "list":
            return 0, f"{cell} must use List validation"
        actual = norm(dv.formula1)
        accepted = {norm(x) for x in accepted_sources}
        if actual in accepted:
            return 1, "Correct"
        return 0, f"Incorrect source on {cell}: {dv.formula1}"


def norm_sheet(value):
    """Normalize Excel sheet names."""
    s = str(value).strip()
    s = s.strip("'")
    return s.upper()


def norm_ref(value):
    """
    Normalize a cell/range reference.

    Examples:
        $G$4:$G$8 -> G4:G8
        'NAMED MANAGER'!$G$4:$G$8
            -> 'NAMED MANAGER'!G4:G8
    """
    if value is None:
        return ""

    s = str(value).strip()
    s = s.replace("$", "")
    s = re.sub(r"\s+", "", s)

    return s.upper()


def normalize_defined_name_ref(attr_text):
    """
    Normalize Defined Name destination.

    Example:
        'NAMED MANAGER'!$G$4:$G$8
        -> 'NAMED MANAGER'!G4:G8
    """
    if not attr_text:
        return ""

    s = str(attr_text).strip()
    s = s.replace("$", "")
    s = re.sub(r"\s+", "", s)
    return s.upper()


def quote_sheet_name(sheet_name):
    """Return Excel-style quoted sheet name."""
    return "'" + sheet_name.replace("'", "''") + "'"


def expected_name_ref(sheet_name, cell_range):
    return norm_ref(
        f"{quote_sheet_name(sheet_name)}!{cell_range}"
    )


def get_defined_names(wb):
    """
    Return workbook defined names as:
        {NAME: [destination refs]}
    """
    names = {}

    for dn in wb.defined_names.values():
        name = str(dn.name).strip().upper()

        try:
            destinations = list(
                dn.destinations
            )
        except Exception:
            destinations = []

        refs = []

        for sheet_name, cell_range in destinations:
            refs.append(
                expected_name_ref(
                    sheet_name,
                    cell_range
                )
            )

        names[name] = refs

    return names


def find_defined_name(wb, name):
    """
    Find a workbook-level Defined Name.
    """
    target = name.strip().upper()

    for dn in wb.defined_names.values():
        if str(dn.name).strip().upper() == target:
            return dn

    return None


def normalize_list_items(text):
    """
    Normalize list-validation values.

    Accepts comma or semicolon separators.
    """
    if text is None:
        return []

    s = str(text).strip()

    if s.startswith("="):
        s = s[1:]

    s = s.strip('"').strip("'")

    parts = re.split(r"[,;]", s)

    return [
        p.strip().upper()
        for p in parts
        if p.strip()
    ]


def get_validation_for_cell(ws, cell_address):
    """
    Return the first DataValidation object covering cell_address.
    """
    target = ws[cell_address].coordinate

    for dv in ws.data_validations.dataValidation:
        sqref = str(dv.sqref)

        for part in sqref.split():
            if validation_range_contains(
                ws,
                part,
                target
            ):
                return dv

    return None


def validation_range_contains(ws, range_text, target_cell):
    """
    Check whether target cell belongs to an sqref range.
    Supports single cells and normal rectangular ranges.
    """
    try:
        min_col, min_row, max_col, max_row = (
            openpyxl.utils.cell.range_boundaries(
                range_text.replace("$", "")
            )
        )

        target_col, target_row = (
            openpyxl.utils.cell.coordinate_to_tuple(
                target_cell
            )[1],
            openpyxl.utils.cell.coordinate_to_tuple(
                target_cell
            )[0]
        )

        # coordinate_to_tuple returns (row, col)
        target_row, target_col = (
            openpyxl.utils.cell.coordinate_to_tuple(
                target_cell
            )
        )

        return (
            min_row <= target_row <= max_row
            and
            min_col <= target_col <= max_col
        )

    except Exception:
        return norm_ref(range_text) == norm_ref(
            target_cell
        )


def normalize_formula_text(value):
    if value is None:
        return ""

    s = str(value).strip().upper()
    s = s.replace("$", "")
    s = re.sub(r"\s+", "", s)

    if s.startswith("="):
        s = s[1:]

    return s


def is_same_formula(value, accepted):
    return (
        normalize_formula_text(value)
        ==
        normalize_formula_text(accepted)
    )


# ============================================================
# RESULT HELPERS
# ============================================================

def new_check():
    return {
        "marks": 0,
        "status": "WRONG",
        "checks": {},
        "issues": []
    }


def finalize_check(result, mark):
    if result["checks"] and all(
        value == 1
        for value in result["checks"].values()
    ):
        result["marks"] = mark
        result["status"] = "CORRECT"
    else:
        result["marks"] = 0
        result["status"] = "WRONG"

    return result


def add_issue_for_failed_checks(result):
    for name, value in result["checks"].items():
        if value == 0:
            result["issues"].append(
                name + " is incorrect."
            )


# ============================================================
# MODULE 1 - NAMED MANAGER
#
# Expected:
# ProductList  -> NAMED MANAGER!G4:G8
# PriceList    -> NAMED MANAGER!H4:H8
# CategoryList -> NAMED MANAGER!I4:I8
#
# 2.5 marks total
# 0.833333... each
# ============================================================

NAMED_MANAGER_CHECKS = {
    "ProductList": (
        "NAMED MANAGER",
        "G4:G8"
    ),

    "PriceList": (
        "NAMED MANAGER",
        "H4:H8"
    ),

    "CategoryList": (
        "NAMED MANAGER",
        "I4:I8"
    ),
}


def check_named_range(
    wb,
    name,
    expected_sheet,
    expected_range
):
    result = new_check()

    dn = find_defined_name(
        wb,
        name
    )

    if dn is None:
        result["issues"].append(
            f"Named range '{name}' not found."
        )
        return result

    try:
        destinations = list(
            dn.destinations
        )
    except Exception:
        destinations = []

    expected = expected_name_ref(
        expected_sheet,
        expected_range
    )

    actual_refs = set()

    for sheet_name, cell_range in destinations:
        actual_refs.add(
            expected_name_ref(
                sheet_name,
                cell_range
            )
        )

    result["checks"]["name_exists"] = 1

    result["checks"]["correct_reference"] = int(
        expected in actual_refs
    )

    if result["checks"]["correct_reference"] == 0:
        result["issues"].append(
            f"{name} should refer to "
            f"{expected_sheet}!{expected_range}."
        )

    if all(
        value == 1
        for value in result["checks"].values()
    ):
        result["marks"] = (
            NAMED_MANAGER_MARKS
            / len(NAMED_MANAGER_CHECKS)
        )
        result["status"] = "CORRECT"

    return result


# ============================================================
# MODULE 2 - BASIC DROPDOWN
#
# C5:
#   Apple, Mango, Banana, Orange
#
# C7:
#   =ProductList
#
# 2.5 marks total
# 1.25 each
# ============================================================

def check_basic_c5(ws):
    result = new_check()

    dv = get_validation_for_cell(
        ws,
        "C5"
    )

    if dv is None:
        result["issues"].append(
            "No Data Validation found on C5."
        )
        return result

    result["checks"]["validation_exists"] = 1

    result["checks"]["type_list"] = int(
        str(dv.type).lower() == "list"
    )

    actual_items = normalize_list_items(
        dv.formula1
    )

    expected_items = [
        "APPLE",
        "MANGO",
        "BANANA",
        "ORANGE"
    ]

    result["checks"]["correct_list"] = int(
        actual_items == expected_items
    )

    add_issue_for_failed_checks(result)

    if all(
        value == 1
        for value in result["checks"].values()
    ):
        result["marks"] = BASIC_DROPDOWN_MARKS / 2
        result["status"] = "CORRECT"

    return result


def check_basic_c7(ws, wb):
    result = new_check()

    dv = get_validation_for_cell(
        ws,
        "C7"
    )

    if dv is None:
        result["issues"].append(
            "No Data Validation found on C7."
        )
        return result

    result["checks"]["validation_exists"] = 1

    result["checks"]["type_list"] = int(
        str(dv.type).lower() == "list"
    )

    formula = normalize_formula_text(
        dv.formula1
    )

    result["checks"]["uses_ProductList"] = int(
        formula == "PRODUCTLIST"
    )

    add_issue_for_failed_checks(result)

    if all(
        value == 1
        for value in result["checks"].values()
    ):
        result["marks"] = BASIC_DROPDOWN_MARKS / 2
        result["status"] = "CORRECT"

    return result


# ============================================================
# MODULE 3 - ADVANCED / DEPENDENT DROPDOWNS
#
# Named ranges:
# Electronics -> DROPDOWN ADVANCED!G4:G5
# Furniture   -> DROPDOWN ADVANCED!H4:H5
#
# C10:
#   Electronics, Furniture
#
# D10:
#   =INDIRECT(C10)
#
# 2.5 marks total
# 4 subchecks
# 0.625 each
# ============================================================

ADVANCED_NAMED_RANGES = {
    "Electronics": (
        "DROPDOWN ADVANCED",
        "G4:G5"
    ),

    "Furniture": (
        "DROPDOWN ADVANCED",
        "H4:H5"
    ),
}


def check_advanced_named_range(
    wb,
    name,
    expected_sheet,
    expected_range
):
    result = new_check()

    if name.lower() == "category":
        accepted_refs = [
            "'DROPDOWN ADVANCED'!$G$3:$H$3",
            "'DROPDOWN ADVANCED'!G3:H3",
            "G3:H3",
            "$G$3:$H$3"
        ]
    elif name.lower() == "electronics":
        accepted_refs = [
            "'DROPDOWN ADVANCED'!$G$4:$G$5",
            "'DROPDOWN ADVANCED'!G4:G5",
            "G4:G5",
            "$G$4:$G$5"
        ]
    elif name.lower() == "furniture":
        accepted_refs = [
            "'DROPDOWN ADVANCED'!$H$4:$H$5",
            "'DROPDOWN ADVANCED'!H4:H5",
            "H4:H5",
            "$H$4:$H$5"
        ]
    else:
        accepted_refs = [
            f"'{expected_sheet}'!${expected_range}$",
            f"'{expected_sheet}'!{expected_range}",
            expected_range
        ]

    dn = wb.defined_names.get(name) or find_defined_name(wb, name)

    if dn is None:
        result["issues"].append(
            f"Named range '{name}' not found."
        )
        return result

    actual = norm(dn.attr_text)
    accepted = {
        norm(x)
        for x in accepted_refs
    }

    correct = int(actual in accepted)

    result["checks"]["name_exists"] = 1
    result["checks"]["correct_reference"] = correct

    if correct == 0:
        result["issues"].append(
            f"{name} refers to incorrect range: {dn.attr_text}"
        )

    add_issue_for_failed_checks(result)

    if all(
        value == 1
        for value in result["checks"].values()
    ):
        result["marks"] = (
            ADVANCED_DROPDOWN_MARKS / 4
        )
        result["status"] = "CORRECT"

    return result


def check_advanced_c10(ws):
    result = new_check()

    accepted_sources = [
        "=Category",
        "=CATEGORY",
        "=CategoryList",
        "=CATEGORYLIST",
        "=Categories",
        "=CATEGORIES",
        "=$G$3:$H$3",
        "$G$3:$H$3",
        "=G3:H3",
        "G3:H3",
        '"Electronics,Furniture"',
        "Electronics,Furniture"
    ]

    code, msg = check_list_dv(ws, "C10", accepted_sources)
    dv = get_dv(ws, "C10")

    result["checks"]["validation_exists"] = 1 if dv is not None else 0
    result["checks"]["type_list"] = 1 if dv is not None and str(dv.type).lower() == "list" else 0
    result["checks"]["correct_categories"] = code

    if code == 0:
        result["issues"].append(msg)

    add_issue_for_failed_checks(result)

    if all(
        value == 1
        for value in result["checks"].values()
    ):
        result["marks"] = (
            ADVANCED_DROPDOWN_MARKS / 4
        )
        result["status"] = "CORRECT"

    return result


def check_advanced_c11(ws):
    result = new_check()
    code, msg = check_c11(ws)

    result["checks"]["dependent_formula"] = code

    if code == 0:
        result["issues"].append(msg)

    add_issue_for_failed_checks(result)

    if all(
        value == 1
        for value in result["checks"].values()
    ):
        result["marks"] = (
            ADVANCED_DROPDOWN_MARKS / 4
        )
        result["status"] = "CORRECT"

    return result


# ============================================================
# MODULE 4 - DATA VALIDATION SKILL 2
#
# C5:
#   Whole Number between 10 and 100
#
# C7:
#   Date between 2024-01-01 and 2024-12-31
#
# C9:
#   Text Length exactly 5
#
# 2.5 marks total
# 0.833333... each
# ============================================================

def check_operator_between(dv):
    if dv is None:
        return 0

    operator = str(dv.operator or "").strip().lower()

    if operator in ["between", ""] or (not operator and getattr(dv, 'formula1', None) is not None and getattr(dv, 'formula2', None) is not None):
        return 1

    return 0


def check_whole_number_c5(ws):
    result = new_check()

    dv = get_validation_for_cell(
        ws,
        "C5"
    )

    if dv is None:
        result["issues"].append(
            "No Data Validation found on C5."
        )
        return result

    result["checks"]["validation_exists"] = 1

    result["checks"]["type_whole"] = int(
        str(dv.type or "").strip().lower() == "whole"
    )

    result["checks"]["operator_between"] = check_operator_between(dv)

    result["checks"]["minimum_10"] = int(
        norm(dv.formula1) == "10"
    )

    result["checks"]["maximum_100"] = int(
        norm(dv.formula2) == "100"
    )

    add_issue_for_failed_checks(result)

    if all(
        value == 1
        for value in result["checks"].values()
    ):
        result["marks"] = (
            DATA_VALIDATION_MARKS / 3
        )
        result["status"] = "CORRECT"

    return result


def date_formula_matches(
    value,
    expected_year,
    expected_month,
    expected_day
):
    """
    Accept:
      DATE(2024,1,1)
      1/1/2024
      45292 (Excel serial)
    """

    s = normalize_formula_text(value)

    expected_date = datetime(
        expected_year,
        expected_month,
        expected_day
    )

    # DATE(2024,1,1)
    if s == (
        f"DATE({expected_year},"
        f"{expected_month},"
        f"{expected_day})"
    ):
        return True

    # Numeric Excel serial
    try:
        if float(s) == (
            expected_date - datetime(1899, 12, 30)
        ).days:
            return True
    except Exception:
        pass

    # Common date text formats
    raw = str(value).strip() if value is not None else ""

    for fmt in [
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%d/%m/%Y",
        "%m-%d-%Y",
        "%d-%m-%Y",
    ]:
        try:
            parsed = datetime.strptime(
                raw,
                fmt
            )
            if parsed.date() == expected_date.date():
                return True
        except Exception:
            pass

    return False


def check_date_c7(ws):
    result = new_check()

    dv = get_validation_for_cell(
        ws,
        "C7"
    )

    if dv is None:
        result["issues"].append(
            "No Data Validation found on C7."
        )
        return result

    result["checks"]["validation_exists"] = 1

    result["checks"]["type_date"] = int(
        str(dv.type).lower() == "date"
    )

    result["checks"]["operator_between"] = check_operator_between(dv)

    result["checks"]["start_date"] = int(
        date_formula_matches(
            dv.formula1,
            2024,
            1,
            1
        )
    )

    result["checks"]["end_date"] = int(
        date_formula_matches(
            dv.formula2,
            2024,
            12,
            31
        )
    )

    add_issue_for_failed_checks(result)

    if all(
        value == 1
        for value in result["checks"].values()
    ):
        result["marks"] = (
            DATA_VALIDATION_MARKS / 3
        )
        result["status"] = "CORRECT"

    return result


def check_text_length_c9(ws):
    result = new_check()

    dv = get_validation_for_cell(
        ws,
        "C9"
    )

    if dv is None:
        result["issues"].append(
            "No Data Validation found on C9."
        )
        return result

    result["checks"]["validation_exists"] = 1

    result["checks"]["type_textLength"] = int(
        str(dv.type).lower() == "textLength".lower()
    )

    result["checks"]["operator_equal"] = int(
        str(dv.operator).lower() == "equal"
    )

    result["checks"]["length_5"] = int(
        norm(dv.formula1) == "5"
    )

    add_issue_for_failed_checks(result)

    if all(
        value == 1
        for value in result["checks"].values()
    ):
        result["marks"] = (
            DATA_VALIDATION_MARKS / 3
        )
        result["status"] = "CORRECT"

    return result


# ============================================================
# CHECK ONE STUDENT FILE
# ============================================================

def check_student_file(file_path):

    result = {
        "Student": file_path.stem,
        "Total": 0,
        "Percentage": 0,
        "Modules": {},
        "Checks": {},
        "Error": ""
    }

    try:

        keep_vba = (
            file_path.suffix.lower()
            == ".xlsm"
        )

        wb = openpyxl.load_workbook(
            file_path,
            data_only=False,
            keep_vba=keep_vba
        )

        total = 0

        # ----------------------------------------------------
        # MODULE 1 - NAMED MANAGER
        # ----------------------------------------------------

        module_marks = 0

        for name, (
            sheet_name,
            cell_range
        ) in NAMED_MANAGER_CHECKS.items():

            check = check_named_range(
                wb,
                name,
                sheet_name,
                cell_range
            )

            result["Checks"][
                f"Named Manager - {name}"
            ] = check

            module_marks += check["marks"]

        result["Modules"][
            "Named Manager"
        ] = round(
            module_marks,
            4
        )

        total += module_marks

        # ----------------------------------------------------
        # MODULE 2 - BASIC DROPDOWN
        # ----------------------------------------------------

        basic_marks = 0

        if "DROPDOWN BASIC" not in wb.sheetnames:

            result["Checks"][
                "Basic Dropdown - C5"
            ] = {
                "marks": 0,
                "status": "WRONG",
                "checks": {},
                "issues": [
                    "Sheet 'DROPDOWN BASIC' not found."
                ]
            }

            result["Checks"][
                "Basic Dropdown - C7"
            ] = {
                "marks": 0,
                "status": "WRONG",
                "checks": {},
                "issues": [
                    "Sheet 'DROPDOWN BASIC' not found."
                ]
            }

        else:

            ws = wb["DROPDOWN BASIC"]

            check = check_basic_c5(ws)
            result["Checks"][
                "Basic Dropdown - C5"
            ] = check
            basic_marks += check["marks"]

            check = check_basic_c7(
                ws,
                wb
            )
            result["Checks"][
                "Basic Dropdown - C7"
            ] = check
            basic_marks += check["marks"]

        result["Modules"][
            "Basic Dropdown"
        ] = round(
            basic_marks,
            4
        )

        total += basic_marks

        # ----------------------------------------------------
        # MODULE 3 - ADVANCED DROPDOWN
        # ----------------------------------------------------

        advanced_marks = 0

        for name, (
            sheet_name,
            cell_range
        ) in ADVANCED_NAMED_RANGES.items():

            check = check_advanced_named_range(
                wb,
                name,
                sheet_name,
                cell_range
            )

            result["Checks"][
                f"Advanced - Named Range - {name}"
            ] = check

            advanced_marks += check["marks"]

        if "DROPDOWN ADVANCED" not in wb.sheetnames:

            for label in [
                "Advanced - Category C10",
                "Advanced - Dependent C11"
            ]:

                result["Checks"][label] = {
                    "marks": 0,
                    "status": "WRONG",
                    "checks": {},
                    "issues": [
                        "Sheet 'DROPDOWN ADVANCED' not found."
                    ]
                }

        else:

            ws = wb["DROPDOWN ADVANCED"]

            check = check_advanced_c10(ws)
            result["Checks"][
                "Advanced - Category C10"
            ] = check
            advanced_marks += check["marks"]

            check = check_advanced_c11(ws)
            result["Checks"][
                "Advanced - Dependent C11"
            ] = check
            advanced_marks += check["marks"]

        result["Modules"][
            "Advanced Dropdown"
        ] = round(
            advanced_marks,
            4
        )

        total += advanced_marks

        # ----------------------------------------------------
        # MODULE 4 - DATA VALIDATION
        # ----------------------------------------------------

        validation_marks = 0

        if "DATA VALIDATION SKILL 2" not in wb.sheetnames:

            for label in [
                "Validation - Whole Number C5",
                "Validation - Date C7",
                "Validation - Text Length C9"
            ]:

                result["Checks"][label] = {
                    "marks": 0,
                    "status": "WRONG",
                    "checks": {},
                    "issues": [
                        "Sheet 'DATA VALIDATION SKILL 2' not found."
                    ]
                }

        else:

            ws = wb[
                "DATA VALIDATION SKILL 2"
            ]

            check = check_whole_number_c5(ws)
            result["Checks"][
                "Validation - Whole Number C5"
            ] = check
            validation_marks += check["marks"]

            check = check_date_c7(ws)
            result["Checks"][
                "Validation - Date C7"
            ] = check
            validation_marks += check["marks"]

            check = check_text_length_c9(ws)
            result["Checks"][
                "Validation - Text Length C9"
            ] = check
            validation_marks += check["marks"]

        result["Modules"][
            "Data Validation"
        ] = round(
            validation_marks,
            4
        )

        total += validation_marks

        # ----------------------------------------------------
        # TOTAL
        # ----------------------------------------------------

        result["Total"] = round(
            total,
            2
        )

        result["Percentage"] = round(
            (total / 10) * 100,
            2
        )

    except Exception as e:

        result["Error"] = str(e)

    return result


# ============================================================
# CREATE EXCEL REPORT
# ============================================================

def create_report(results):

    RESULTS_DIR.mkdir(
        exist_ok=True
    )

    output = (
        RESULTS_DIR
        / "Excel_Skill_2_Marks_Report.xlsx"
    )

    wb = openpyxl.Workbook()

    # ========================================================
    # MARKS SHEET
    # ========================================================

    ws = wb.active
    ws.title = "Marks"

    module_names = [
        "Named Manager",
        "Basic Dropdown",
        "Advanced Dropdown",
        "Data Validation"
    ]

    headers = (
        ["Student"]
        + module_names
        + ["Total / 10", "Percentage", "Error"]
    )

    ws.append(headers)

    for result in results:

        row = [
            result["Student"]
        ]

        for module in module_names:
            row.append(
                result["Modules"].get(
                    module,
                    0
                )
            )

        row.extend([
            result["Total"],
            result["Percentage"],
            result["Error"]
        ])

        ws.append(row)

    # ========================================================
    # DETAILED CHECKING
    # ========================================================

    detail = wb.create_sheet(
        "Detailed Checking"
    )

    detail.append([
        "Student",
        "Check",
        "0_1",
        "Marks",
        "Status",
        "Checks",
        "Issues"
    ])

    for result in results:

        for check_name, check in result[
            "Checks"
        ].items():

            detail.append([

                result["Student"],

                check_name,

                1 if check.get(
                    "status"
                ) == "CORRECT" else 0,

                check.get(
                    "marks",
                    0
                ),

                check.get(
                    "status"
                ),

                json.dumps(
                    check.get(
                        "checks",
                        {}
                    )
                ),

                " | ".join(
                    check.get(
                        "issues",
                        []
                    )
                )
            ])

    # ========================================================
    # VALIDATION / NAME ANALYSIS
    # ========================================================

    analysis = wb.create_sheet(
        "Formula & Validation Analysis"
    )

    analysis.append([
        "Student",
        "Check",
        "Status",
        "Marks",
        "0_1",
        "Checks",
        "Issues"
    ])

    for result in results:

        for check_name, check in result[
            "Checks"
        ].items():

            analysis.append([

                result["Student"],

                check_name,

                check.get(
                    "status"
                ),

                check.get(
                    "marks",
                    0
                ),

                1 if check.get(
                    "status"
                ) == "CORRECT" else 0,

                json.dumps(
                    check.get(
                        "checks",
                        {}
                    )
                ),

                " | ".join(
                    check.get(
                        "issues",
                        []
                    )
                )
            ])

    # ========================================================
    # FORMATTING
    # ========================================================

    for sheet in wb.worksheets:

        sheet.freeze_panes = "A2"

        for column in sheet.columns:

            max_length = 0

            letter = (
                column[0]
                .column_letter
            )

            for cell in column:

                try:
                    max_length = max(
                        max_length,
                        len(
                            str(
                                cell.value
                            )
                        )
                    )
                except Exception:
                    pass

            sheet.column_dimensions[
                letter
            ].width = min(
                max_length + 2,
                60
            )

    wb.save(output)

    return output


# ============================================================
# MAIN
# ============================================================

def main():

    print("=" * 70)
    print(
        "       EXCEL SKILL 2 - AUTO CHECKER"
    )
    print(
        "       DATA VALIDATION & NAME MANAGER"
    )
    print("=" * 70)

    SUBMISSIONS_DIR.mkdir(
        exist_ok=True
    )

    RESULTS_DIR.mkdir(
        exist_ok=True
    )

    files = []

    for extension in [
        "*.xlsx",
        "*.xlsm"
    ]:

        files.extend(
            SUBMISSIONS_DIR.glob(
                extension
            )
        )

    if not files:

        print(
            "\nNo student Excel files found."
        )

        print(
            f"Put all student files inside:\n"
            f"{SUBMISSIONS_DIR}"
        )

        return

    files = sorted(files)

    print(
        f"\nFound {len(files)} student files.\n"
    )

    results = []

    for index, file_path in enumerate(
        files,
        start=1
    ):

        print(
            f"[{index}/{len(files)}] "
            f"Checking {file_path.name} ..."
        )

        result = check_student_file(
            file_path
        )

        results.append(
            result
        )

        print(
            f"    Score: "
            f"{result['Total']}/10 "
            f"({result['Percentage']}%)"
        )

        if result["Error"]:
            print(
                f"    ERROR: "
                f"{result['Error']}"
            )

    report = create_report(
        results
    )

    print(
        "\n" + "=" * 70
    )
    print(
        "CHECKING COMPLETE"
    )
    print("=" * 70)

    print(
        f"Students checked : "
        f"{len(results)}"
    )

    print(
        f"Report           : "
        f"{report}"
    )


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    main()
