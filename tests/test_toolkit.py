import unittest
import os
import sys

# Add scripts dir to path dynamically
SCRIPTS_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "scripts"))
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

from engine.formula_parser import shift_formula_references, col_to_num, num_to_col
from engine.sniffer import detect_excel_flavor

class TestExcelToolkit(unittest.TestCase):

    def test_col_conversions(self):
        self.assertEqual(col_to_num('A'), 1)
        self.assertEqual(col_to_num('K'), 11)
        self.assertEqual(col_to_num('Z'), 26)
        self.assertEqual(col_to_num('AA'), 27)
        self.assertEqual(num_to_col(1), 'A')
        self.assertEqual(num_to_col(11), 'K')
        self.assertEqual(num_to_col(26), 'Z')
        self.assertEqual(num_to_col(27), 'AA')

    def test_formula_shifting(self):
        # Insert column at K (11)
        # Columns before 11 remain unchanged
        self.assertEqual(shift_formula_references("=A2+B2", 11, 1), "=A2+B2")
        self.assertEqual(shift_formula_references("=J2*10", 11, 1), "=J2*10")
        
        # Columns >= 11 shift by 1 (L->M, N->O, T->U)
        self.assertEqual(shift_formula_references("=(L2-N2)/L2", 11, 1), "=(M2-O2)/M2")
        self.assertEqual(shift_formula_references("=N2*1.1", 11, 1), "=O2*1.1")
        self.assertEqual(shift_formula_references("=(L2-T2)/L2", 11, 1), "=(M2-U2)/M2")
        
        # Complex formula with IF and SUM
        complex_f = "=IF(L2>0, SUM(N2:T2)/L2, 0)"
        expected_f = "=IF(M2>0, SUM(O2:U2)/M2, 0)"
        self.assertEqual(shift_formula_references(complex_f, 11, 1), expected_f)

    def test_sniffer(self):
        wps_file = r"C:\Users\q1894\Desktop\货盘已核查有效链接.xlsx"
        if os.path.exists(wps_file):
            flavor = detect_excel_flavor(wps_file)
            self.assertEqual(flavor, 'wps')

if __name__ == '__main__':
    unittest.main()
