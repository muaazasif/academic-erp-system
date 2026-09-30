import unittest
import openpyxl
from excel_assignment import check_q1_s1, check_q2_s1, check_q3_s1, check_q4_s1

class TestExcelSkill1VLOOKUPExpert(unittest.TestCase):

    def test_q1_expert_variations(self):
        # Q1: Lookup "E003", start col A, return col 3 (Department)
        valid_formulas = [
            '=VLOOKUP("E003", A4:C13, 3, FALSE)',
            '=VLOOKUP("E003", A4:E13, 3, 0)',
            '=VLOOKUP("E003", A:C, 3, FALSE)',
            '=VLOOKUP("E003", $A$4:$E$13, 3, 0)',
            '=VLOOKUP("E003", A3:C13, 3, TRUE)',
            '=VLOOKUP("E003", A3:E13, 3, 1)',
            '=VLOOKUP("E003", A:E, 3, FALSE)',
            '=VLOOKUP("E003", A4:C100, 3, 0)',
            '=VLOOKUP("E003", A4:E100, 3, FALSE)',
            '=VLOOKUP("E003", $A$4:$C$13, 3, 0)',
            '=VLOOKUP("E003", $A$4:$E$13, 3, FALSE)',
            '=VLOOKUP("E003", A$4:C$13, 3, 0)',
            '=VLOOKUP($B$20, A4:C13, 3, FALSE)', # cell reference lookup
        ]
        for f in valid_formulas:
            res = check_q1_s1(f)
            self.assertEqual(res["status"], "CORRECT", f"Failed for Q1 formula: {f}")

    def test_q2_expert_variations(self):
        # Q2: Lookup "Sara Khan", start col B, return col 4 (Salary)
        valid_formulas = [
            '=VLOOKUP("Sara Khan", B4:E13, 4, FALSE)',
            '=VLOOKUP("Sara Khan", B4:E13, 4, 0)',
            '=VLOOKUP("Sara Khan", B:E, 4, FALSE)',
            '=VLOOKUP("Sara Khan", $B$4:$E$13, 4, 0)',
            '=VLOOKUP("Sara Khan", B3:E13, 4, TRUE)',
            '=VLOOKUP("Sara Khan", B4:E100, 4, 1)',
            '=VLOOKUP($B$20, B4:E13, 4, FALSE)',
        ]
        for f in valid_formulas:
            res = check_q2_s1(f)
            self.assertEqual(res["status"], "CORRECT", f"Failed for Q2 formula: {f}")

        # Q2 should reject wrong start col (e.g. A:E)
        invalid_formulas = [
            '=VLOOKUP("Sara Khan", A4:E13, 4, FALSE)',
            '=VLOOKUP("Sara Khan", A:E, 5, FALSE)',
        ]
        for f in invalid_formulas:
            res = check_q2_s1(f)
            self.assertEqual(res["status"], "WRONG", f"Should have failed for Q2 formula: {f}")

    def test_q3_expert_variations(self):
        # Q3: Lookup "E007", start col A, return col 4 (City)
        valid_formulas = [
            '=VLOOKUP("E007", A4:D13, 4, FALSE)',
            '=VLOOKUP("E007", A4:E13, 4, 0)',
            '=VLOOKUP("E007", A:D, 4, FALSE)',
            '=VLOOKUP("E007", $A$4:$E$13, 4, 0)',
            '=VLOOKUP("E007", A3:D13, 4, TRUE)',
            '=VLOOKUP("E007", A:E, 4, 1)',
        ]
        for f in valid_formulas:
            res = check_q3_s1(f)
            self.assertEqual(res["status"], "CORRECT", f"Failed for Q3 formula: {f}")

    def test_q4_expert_variations(self):
        # Q4: Lookup "E010", start col A, return col 2 (Name)
        valid_formulas = [
            '=VLOOKUP("E010", A4:B13, 2, FALSE)',
            '=VLOOKUP("E010", A4:E13, 2, 0)',
            '=VLOOKUP("E010", A:B, 2, FALSE)',
            '=VLOOKUP("E010", $A$4:$E$13, 2, 0)',
            '=VLOOKUP("E010", A3:B13, 2, TRUE)',
            '=VLOOKUP("E010", A:E, 2, 1)',
        ]
        for f in valid_formulas:
            res = check_q4_s1(f)
            self.assertEqual(res["status"], "CORRECT", f"Failed for Q4 formula: {f}")

if __name__ == '__main__':
    unittest.main()
