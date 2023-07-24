import unittest
import pandas as pd

class TestExcelFiles(unittest.TestCase):
    def test_DEFAULT_(self):
    # Excel files you wish to validate along with their names
        filenames = [
            ("TestCasesDefault\Time Detail_July152022.xlsx", "time weston.xlsx"),
            ("TestCasesDefault\Time Detail_July152022.xlsx", "time weston minutes.xlsx")
        ]

        # Values you want to search for in each file
        name_value = "Anderson, Kasey "
        glcode_value = "3050-3001-31233"

        # List to store the details of the rows that do not meet the criteria.
        failed_rows = []

        for filename, excel_name in filenames:
            # Read the Excel file
            df = pd.read_excel(filename)

            # Filter by the value in the "NAME" column
            filtered_df = df[df["NAME"] == name_value]

            # Get the rows that do not meet the validation criterion
            incorrect_rows = filtered_df[filtered_df["GLCODE"] != glcode_value]

            # If there are incorrect rows, add the details to the list of failed_rows
            if not incorrect_rows.empty:
                for index, row in incorrect_rows.iterrows():
                    failed_rows.append((excel_name, index + 2, row["GLCODE"]))

        # Check if any row did not meet the criterion
        if failed_rows:
            # Display the message with the details of the incorrect rows
            message = "Problems were found in the following records:\n"
            for excel_name, row_num, glcode in failed_rows:
                message += f"File: {excel_name}, Row {row_num}: GLCODE={glcode} does not match the expected value.\n"
            self.fail(message)
        
        print(".TEST 7 DEFAULT CORRECT: Kasey's GLCODE contains - and is 3050-3001-31233")

if __name__ == '__main__':
    unittest.main()
