import unittest
import pandas as pd
import numpy as np

class TestExcel(unittest.TestCase):

    def test_UKGK_1(self):

        # Loads the Excel file into a DataFram
        df = pd.read_excel('TestCasesUKGSimplified/UKG Simplified Empty.xlsx', header=None)

        # Gets the number of rows with data beyond the headers
        num_data_rows = len(df) - 1 

        # Check if there are additional rows with data
        if num_data_rows > 0:

            self.fail("There are {} additional rows with data in the Excel file".format(num_data_rows))
        
        print(".TEST 1 UKGSimplified CORRECT: The UKG Kronos Empty.xlsx file not contains additional rows to the header")

    def test_UKGK_2(self):

        #Descrition: Check last shift is present
        #File: Hannibal

            # Upload Excel file
            excel_file = 'TestCasesAPI/Hannibal 4.15.23 SCHED.xlsx'
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
                self.assertEqual(len(filtered_df), 1, 'Multiple matches found in Excel file.')

            print(".TEST 6 API CORRECT: Checked that the last line of the file TestCasesAPI/Hannibal 4.15.23 SCHED.xlsx is still for WYCOFF, JENNA with the same data")