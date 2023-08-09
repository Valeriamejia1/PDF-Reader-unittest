import unittest
import pandas as pd

class ExcelTest(unittest.TestCase):
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
        #Files: Blinder1

    def test_Paylocity_1(self):
        self.compare_excel_files("TestCasesPaylocity/Binder1 ORIG.xlsx", "OUTPUT Paylocity/Binder1.xlsx")
        print("TEST 1 Paylocity CORRECT: Blinder1.xlsx data match the original version")


if __name__ == '__main__':
    unittest.main()
