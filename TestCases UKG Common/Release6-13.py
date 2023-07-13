import unittest
import pandas as pd

class ExcelTest(unittest.TestCase):
    def testCase1(self):
        # Description: Check last shift is present
        # File: Martin ppe
        # Upload Excel file
        excel_file = 'Martin ppe 4.22.23.xlsx'
        df = pd.read_excel(excel_file, sheet_name='Sheet1')

        # Specify the search criteria
        name = 'WARREN, DANIELLE'
        date = '04/22/2023'
        paycode = 'Regular'
        hours = 11.00

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
                print("The data is correct.")

            print("TEST 1 CORRECT:Data for WARREN, DANIELLE was found and is correct.")
    
    def compare_excel_files(self, original_file, new_file):
        #Descrition: Check Exe has the same data as last commit
        #Files: All Files

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


    def testMartinppe(self):
        self.compare_excel_files("Martin ppe 4.22.23 ORIG.xlsx", "Martin ppe 4.22.23.xlsx")
        print("Martin ppe 4.22.23.xlsx data match the original version")

    def testMartinB(self):
        self.compare_excel_files("martin b ORIG.xlsx", "martin b.xlsx")
        print("Martin ppe 4.22.23.xlsx data match the original version")

    def testTimeMartin(self):
        self.compare_excel_files("time martin a ORIG.xlsx", "time martin a.xlsx")
        print("Martin ppe 4.22.23.xlsx data match the original version")

if __name__ == '__main__':
    unittest.main()
