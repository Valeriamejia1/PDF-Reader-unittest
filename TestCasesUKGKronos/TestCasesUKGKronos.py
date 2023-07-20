import unittest
import pandas as pd

class TestExcel(unittest.TestCase):

    def test_UKGK_1(self):

        # Loads the Excel file into a DataFram
        df = pd.read_excel('TestCasesUKGKronos/UKG Kronos empty.xlsx', header=None)

        # Gets the number of rows with data beyond the headers
        num_data_rows = len(df) - 1 

        # Check if there are additional rows with data
        if num_data_rows > 0:

            self.fail("There are {} additional rows with data in the Excel file".format(num_data_rows))
        
        print(".TEST 1 UKGKronos CORRECT: The UKG Kronos Empty.xlsx file not contains additional rows to the header")

    def test_UKGK_2(self):

        #Descrition: Check last shift is present
        #File: Hannibal

            # Upload Excel file
            excel_file = 'TestCasesUKGKronos\Qualivis Time report PPE 062423.xlsx'
            df = pd.read_excel(excel_file)
       
            # Specify the search criteria
            name = 'WUISCHPARD, DAVID'
            date = '06/21/2023'
            hours = 12.25

            # Filter the DataFrame based on the following criteria
            filtered_df = df[
                (df['NAME'] == name) &
                (df['DATE'] == date) &
                (df['HOURS'] == hours)
            ]

            # Check if matching rows were found
            if filtered_df.empty:
                error_message = f"No matching row was found in the Excel file for the following values:\n\n" \
                                f"Name: {name}\n" \
                                f"Date: {date}\n" \
                                f"Hours: {hours}\n"
                self.fail(error_message)
            else:
                self.assertEqual(len(filtered_df), 1, 'Multiple matches found in Excel file.')

            print(".TEST 2 UKG Kronos CORRECT: Checked that the last line of the file Qualivis Time report PPE 062423.xlsx is still for WUISCHPARD, DAVID with the same data")

    def test_UKGK_6(self):
        # Load the Excel files into pandas DataFrames
        expected_df = pd.read_excel("TestCasesUKGKronos/Qualivis Time report PPE 062423 ORIG.xlsx")
        actual_df = pd.read_excel("TestCasesUKGKronos/Qualivis Time report PPE 062423.xlsx")

        # Fill NaN values in both DataFrames with empty strings
        expected_df = expected_df.fillna('')
        actual_df = actual_df.fillna('')

        # Compare the "NAME" and "Comments" columns
        diff_rows = actual_df[(actual_df['NAME'] != expected_df['NAME']) | (actual_df['Comments'] != expected_df['Comments'])]

        # Check if there are any differences
        if not diff_rows.empty:
            for index, row in diff_rows.iterrows():
                print(f"Row {index + 2}: NAME - Expected: '{expected_df.loc[index, 'NAME']}', Actual: '{row['NAME']}' | "
                      f"Comments - Expected: '{expected_df.loc[index, 'Comments']}', Actual: '{row['Comments']}'")
            raise AssertionError("Data mismatch found between the two Excel files.")
        else:
            print("TEST 6 UKGKronos CORRECT: The 'NAME' and 'Comments' columns in the Excel files are the same.")

    def test_UKGK_8(self):
        # Load the Excel file into a pandas DataFrame
        df = pd.read_excel("TestCasesUKGKronos/Qualivis Time report PPE 062423.xlsx")

        # Check if row 555 contains the expected values
        expected_name = "MCCLAM, REBECCA"
        expected_hours = ""

        row_data = df.loc[553].copy()  # Row index is 0-based, so row 555 is at index 553 in DataFrame

        # Check if the 'NAME' column contains the expected name
        self.assertEqual(row_data['NAME'], expected_name, f"Name MCCLAM, REBECCA does not match the expected value.")

        # Convert NaN values to empty strings in the 'HOURS' column
        if pd.isna(row_data['HOURS']):
            row_data.loc['HOURS'] = ''

        # Check if the 'HOURS' column contains the expected hours (empty sheet)
        self.assertEqual(row_data['HOURS'], expected_hours, f"Hours do not match the expected value 'EMPTY'.")

        print("TEST 8 UKGKronos CORRECT: The last row for MCCLAM, REBECCA is correct, hours empty.")

    def test_UKGK_9(self):
        # Load the Excel file into a pandas DataFrame
        df = pd.read_excel("TestCasesUKGKronos/Qualivis Time report PPE 062423.xlsx")

        # Check if row 731 contains the expected values
        expected_name = "PALMER, NATALIE"
        expected_hours = "9.75"

        row_data = df.loc[729]  # Row index is 0-based, so row 731 is at index 730 in DataFrame

        # Check if the 'NAME' column contains the expected name
        self.assertEqual(row_data['NAME'], expected_name, f"Name PALMER, NATALIE does not match the expected value.")

        # Convert the value in the 'HOURS' column to a string
        actual_hours = str(row_data['HOURS'])

        # Check if the 'HOURS' column contains the expected hours (as a string)
        self.assertEqual(actual_hours, expected_hours, f"Hours do not match the expected value '9.75'.")

        print("TEST 9 UKGKronos CORRECT: The last row for PALMER, NATALIE is correct, hours is 9.75.")


if __name__ == "__main__":
    unittest.main()






