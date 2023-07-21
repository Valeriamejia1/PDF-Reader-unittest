import unittest
import re
import pandas as pd
from openpyxl import load_workbook

class ExcelTestCase(unittest.TestCase):

    def test_DEFAULT_1(self):

        # Loads the Excel file into a DataFram
        df = pd.read_excel('TestCasesDefault/Default Empty.xlsx', header=None)

        # Gets the number of rows with data beyond the headers
        num_data_rows = len(df) - 1 

        # Check if there are additional rows with data
        if num_data_rows > 0:

            self.fail("There are {} additional rows with data in the Excel file".format(num_data_rows))
        
        print(".TEST 1 DEFAULT CORRECT: The Default Empty.xlsx file not contains additional rows to the header")

    def test_DEFAULT_2(self):
    # Excel files you wish to validate along with their names
        filenames = [
            ("TestCasesDefault/time weston.xlsx", "time weston.xlsx"),
            ("TestCasesDefault/time weston minutes.xlsx", "time weston minutes.xlsx")
        ]

        # Values you want to search for in each file
        name_value = "Celestin, Elizabeth"
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
        
        print(".TEST 2 DEFAULT CORRECT: Celestin's GLCODE, Elizabeth contains - and is 3050-3001-31233")

    #Method required for test_DEFAULT_3

    def validate_glcode(self, glcode):
        # Eliminar los guiones "-" y contar los dígitos restantes
        glcode_sin_guiones = glcode.replace('-', '')
        return len(glcode_sin_guiones) == 9 and glcode_sin_guiones.isdigit()

    def test_DEFAULT_3(self):
        files = ["TestCasesDefault/Combined File minutes.xlsx", "TestCasesDefault/Combined File.xlsx"]

        all_incorrect_rows = set()  # We use a set to store the incorrect rows without duplicates

        for file in files:
            xls = pd.ExcelFile(file)
            sheet_name = xls.sheet_names[0]  # We assume that the sheet of interest is the first one.

            # Load the file and convert the "GLCODE" column to "text" format.
            df = pd.read_excel(xls, sheet_name)
            df["GLCODE"] = df["GLCODE"].astype(str)

            # Validate that all cells in the "GLCODE" column contain 9 digits without dashes "-".
            incorrect_rows = []
            for index, glcode in df["GLCODE"].items():
                if not self.validate_glcode(glcode):
                    incorrect_rows.append((file, index + 2, glcode))

            # Add the incorrect rows to the general set
            all_incorrect_rows.update(incorrect_rows)

        # Show failure message with all incorrect rows of all files
        if all_incorrect_rows:
            error_msg = "GLCODE does not contain 9 digits in the following rows and files:\n"
            for file, fila, glcode in all_incorrect_rows:
                error_msg += f"file: {file}, row: {fila}, GLCODE: {glcode}\n"
            self.fail(error_msg)

        print(".TEST 3 DEFAULT CORRECT: All GLCODEs have 9 numeric digits.")

if __name__ == '__main__':

    unittest.main()
