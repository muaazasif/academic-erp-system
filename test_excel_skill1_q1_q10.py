import unittest
import openpyxl
from excel_assignment import check_q3_s1

class TestExcelSkill1EscapedColonAndVariations(unittest.TestCase):

    def test_q3_escaped_colon_and_variations(self):
        formulas = [
            '=VLOOKUP("E007",A4\:E13,4,0)',
            '=VLOOKUP("E007",A4:E13,4,0)',
            '=VLOOKUP("E007",A4:D13,4,FALSE)',
            '=VLOOKUP(B20,A4\:D13,4,FALSE)',
        ]
        for f in formulas:
            res = check_q3_s1(f)
            self.assertEqual(res["status"], "CORRECT", f"Failed for formula: {f}")

if __name__ == '__main__':
    unittest.main()
