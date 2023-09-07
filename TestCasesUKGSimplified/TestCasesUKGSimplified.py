import unittest
import pandas as pd
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.filterwarnings("ignore", category=DeprecationWarning)
import numpy as np

class ExcelTest(unittest.TestCase):

    def test_UKGS_1(self):

        # Loads the Excel file into a DataFram
        df = pd.read_excel('OUTPUT UKGSimplified/UKG Simplified Empty.xlsx', header=None) 

        # Gets the number of rows with data beyond the headers
        num_data_rows = len(df) - 1 

        # Check if there are additional rows with data
        if num_data_rows > 0:

            self.fail("There are {} additional rows with data in the Excel file".format(num_data_rows))
        
        print(".TEST 1 UKGSimplified CORRECT: The UKG Simplified Empty.xlsx file not contains additional rows to the header")

    def test_UKGS_2(self):
        # Description: Check last shift is present
        # File: TimeDetailsSorted_KEVCOL_2023-04-13T093000.546
        # Upload Excel file
        excel_file = 'OUTPUT UKGSimplified/TimeDetailsSorted_KEVCOL.xlsx'
        df = pd.read_excel(excel_file, sheet_name='Sheet1')

        # Specify the search criteria
        name = 'Wright, Casey L'
        date = '04/08/2023'
        paycode = 'Regular'
        hours = 12.20

        # Create a list of tuples containing the expected values and their corresponding column names
        expected_values = [
            (name, 'NAME'),
            (date, 'DATE'),
            (paycode, 'PAYCODE'),
            (hours, 'HOURS')
        ]

        missing_values = []
        for value, column in expected_values:
            if value not in df[column].values:
                missing_values.append(f"{column}: {value}")

        if missing_values:
            error_message = f"The following value(s) did not match in the Excel file:\n\n"
            error_message += "\n".join(missing_values)
            values_found = df.loc[df['NAME'] == name, ['NAME', 'DATE', 'PAYCODE', 'HOURS']]
            self.fail(f"{error_message}\n\nValues found:\n\n{values_found.to_string(index=False)}")
        else:
            filtered_df = df[
                (df['NAME'] == name) &
                (df['DATE'] == date) &
                (df['PAYCODE'] == paycode) &
                (df['HOURS'] == hours)
            ]
            if filtered_df.empty:
                self.fail("No matching row was found in the Excel file.")
            else:
                self.assertEqual(len(filtered_df), 1, 'Multiple matches found in Excel file.')
                print("TEST 2 UKGSimplified CORRECT:Data for WARREN, DANIELLE was found and is correct in file TimeDetailsSorted_KEVCOL.xlsx.")

    #Methods required for test_UKGC_3

    def compare_excel_files(self, original_file, new_file):


        # Loads the original Excel file in a DataFrame
        original_df = pd.read_excel(original_file, sheet_name="Sheet1")

        # Loads the new Excel file in a DataFrame
        new_df = pd.read_excel(new_file, sheet_name="Sheet1")

        # Verify that the DataFrames are equal
        self.assertTrue(
            original_df.equals(new_df),
            self.generate_difference_message(original_df, new_df, original_file, new_file),
        )

    def generate_difference_message(self, original_df, new_df, original_file, new_file):
        original_df = original_df.fillna("NA")
        new_df = new_df.fillna("NA")

        diff_df = original_df != new_df
        diff_indices = diff_df.any(axis=1)

        diff_message = f"Differences found between {original_file} and {new_file}:\n"
        for index, row in diff_df[diff_indices].iterrows():
            diff_message += f"Row {index+2}:\n"
            for column, value in row.items():
                if value:
                    original_value = original_df.at[index, column]
                    new_value = new_df.at[index, column]
                    diff_message += f"  - Column '{column}': Original='{original_value}', New='{new_value}'\n"

        return diff_message

    #Descrition: Check Exe has the same data as last commit
    #Files: TimeDetailsSorted_KEVCOL, Qualvis TimeSheets 2023-06-03
    
    def test_UKGS_3_1(self):
        self.compare_excel_files("TestCasesUKGSimplified/ORIG Files/TimeDetailsSorted_KEVCOL ORIG.xlsx", "OUTPUT UKGSimplified/TimeDetailsSorted_KEVCOL.xlsx")
        print("TEST 3.1 UKGSimplified CORRECT: TimeDetailsSorted_KEVCOL.xlsx data match the original version")

    def test_UKGS_3_2(self):
        self.compare_excel_files("TestCasesUKGSimplified/ORIG Files/Qualvis TimeSheets 2023-06-03 ORIG.xlsx", "OUTPUT UKGSimplified/Qualvis TimeSheets 2023-06-03.xlsx")
        print("TEST 3.2 UKGSimplified CORRECT: Qualvis TimeSheets 2023-06-03.xlsx data match the original version")

    def test_UKGS_4(self):

        #Description: Verified that all shift are with Paycodes as Regular or Different Paycode 
        #File: Qualvis TimeSheets 2023-06-03.xlsx

        # Upload Excel file
        df = pd.read_excel("OUTPUT UKGSimplified/Qualvis TimeSheets 2023-06-03.xlsx")

        # Verify if the column "PAYCODE" exists in the file
        if "PAYCODE" not in df.columns:
            self.fail("No 'PAYCODE' column found in the Excel file.")

        paycode_column = df["PAYCODE"]
        empty_rows = paycode_column[paycode_column.isnull()]
        
        # Check if there are rows without data in the column "PAYCODE".
        if not empty_rows.empty:
            empty_row_indices = empty_rows.index + 2  # Adjust the indexes
            empty_row_numbers = [f"Row(s): {row}" for row in empty_row_indices]
            self.fail(f"The following rows have no data in the 'PAYCODE' column: {', '.join(empty_row_numbers)}")
        else:
            print("TEST 4 UKGSimplified CORRECT: Column 'PAYCODE' does not contain empty data in its rows in file Qualvis TimeSheets 2023-06-03.xlsx.")

    def test_UKGS_5(self):

        #Descrition: Verified shift without hours and start date are showing up in output file 
        #Files: Qualvis TimeSheets 2023-06-03.xlsx

        # Upload Excel file
        df = pd.read_excel("OUTPUT UKGSimplified/Qualvis TimeSheets 2023-06-03.xlsx")

        # Verify if the column "DATE" exists in the file
        if "DATE" not in df.columns:
            self.fail("No 'Date' column found in the Excel file.")

        date_column = df["DATE"]
        empty_rows = date_column[date_column.isnull()]
        
        # Check if there are rows without data in the column "DATE".
        if not empty_rows.empty:
            empty_row_indices = empty_rows.index + 2  # Adjust the indexes
            empty_row_numbers = [f"Row(s): {row}" for row in empty_row_indices]
            self.fail(f"The following rows have no data in the 'DATE' column: {', '.join(empty_row_numbers)}")
        else:
            print("TEST 5 UKGSimplified CORRECT:Column 'DATE' does not contain empty data in its rows in file Qualvis TimeSheets 2023-06-03.xlsx.")

if __name__ == '__main__':
    unittest.main()
