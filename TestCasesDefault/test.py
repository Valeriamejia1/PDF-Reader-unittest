import unittest
import pandas as pd
from dateutil.parser import parse
import warnings

class ExcelTestCase(unittest.TestCase):

    def test_Default_21(self):
        filenames = [
            r"OUTPUT Default\1690808400472_1671940182 .xlsx",
            r"OUTPUT Default\1690808400472_1671940182 Minutes.xlsx"
        ]
        
        warnings.filterwarnings("ignore", category=FutureWarning, message=".*iteritems is deprecated.*")
        
        for filename in filenames:
            df = pd.read_excel(filename, sheet_name='Sheet1', header=0)
            enddtm_column = df['ENDDTM']
            
            for index, cell_value in enddtm_column.items():
                try:
                    date_object = parse(str(cell_value))
                except ValueError:
                    raise ValueError(f"Invalid date format found in row {index + 2} of file {filename}. Invalid date: {cell_value}")
        
        print("TEST 21 DEFAULT CORRECT: Column Date format is correct in all files")

if __name__ == '__main__':
    unittest.main()
