import unittest
import pandas as pd

class ExcelTest(unittest.TestCase):
    def test_search_in_excel(self):
        # Upload Excel file
        excel_file = 'Hannibal 4.15.23 SCHED.xlsx'
        df = pd.read_excel(excel_file, sheet_name='OutputData')

        # Specify the search criteria
        name = 'WYCOFF, JENNA'
        date = '04/07/2023'
        paycode = 'REG'
        hours = 13.10

        # Filter the DataFrame based on the following criteria
        filtered_df = df[
            (df['NAME'] == name) &
            (df['DATE'] == date) &
            (df['PAYCODE'] == paycode) &
            (df['HOURS'] == hours)
        ]

        # Check if matching rows were found
        if filtered_df.empty:
            error_message = f"No matching row was found in the Excel file for the following values:\n\n" \
                            f"Name: {name}\n" \
                            f"Date: {date}\n" \
                            f"Paycode: {paycode}\n" \
                            f"Hours: {hours}\n"
            self.fail(error_message)
        else:
            self.assertEqual(len(filtered_df), 1, 'Se encontraron múltiples coincidencias en el archivo Excel.')

if __name__ == '__main__':
    unittest.main()
