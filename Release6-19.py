import unittest
import pandas as pd

class ExcelTest(unittest.TestCase):
    def compare_excel_files(self, original_file, new_file):
        try:
            # Loads the original Excel file in a DataFrame
            original_df = pd.read_excel(original_file, sheet_name="OutputData")

            # Loads the new Excel file in a DataFrame
            new_df = pd.read_excel(new_file, sheet_name="OutputData")

            # Verify the number of rows
            self.assertEqual(
                len(new_df),
                len(original_df),
                f"Number of rows is not equal in the comparison of {original_file} and {new_file}",
            )

            # Verify the existence of the column "NAME".
            self.assertIn(
                "NAME",
                new_df.columns,
                f"The column 'NAME' was not found in {new_file}",
            )

            # Verify the values in all cells of the DataFrame
            self.assertTrue(
                new_df.equals(original_df),
                f"{new_file} does not have the same values as {original_file}",
            )

        except AssertionError as e:
            self.fail(str(e))

    def testDawsonKathleen(self):
        self.compare_excel_files("Dawson, Kathleen ORIG.xlsx", "Dawson, Kathleen.xlsx")

    def testMattoxKyle(self):
        self.compare_excel_files("Mattox, Kyle ORIG.xlsx", "Mattox, Kyle.xlsx")

    def testHannibal(self):
        self.compare_excel_files("Hannibal 4.15.23 ORIG.xlsx", "Hannibal 4.15.23.xlsx")


if __name__ == "__main__":
    unittest.main()