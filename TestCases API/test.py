import unittest

import pandas as pd

class ExcelTestCase(unittest.TestCase):

    def test_API_5(self):

        # Loads the Excel file into a DataFram
        df = pd.read_excel('TestCases API\API Empty.xlsx', header=None)

        # Gets the number of rows with data beyond the headers
        num_data_rows = len(df) - 1 

        # Check if there are additional rows with data
        if num_data_rows > 0:

            self.fail("There are {} additional rows with data in the Excel file".format(num_data_rows))
        
        print(".TEST 5 API CORRECT: The API Empty.xlsx file not contains additional rows to the header")

if __name__ == '__main__':

    unittest.main()
