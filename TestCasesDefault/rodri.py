import unittest
import pandas as pd
import re

class ExcelTestCase(unittest.TestCase):

    def test_Default_9(self): 
        # File path and sheet name
        #File: TestCasesDefault\Scripps Approved Kronos we 6.25.22.xlsx
        #Desc: Check column Date format and emptys cells
        file_path = 'TestCasesDefault\\Scripps Approved Kronos we 6.25.22.xlsx'
        sheet_name = 'Sheet1'

        # Read the Excel file and the specified sheet
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            # Check for missing or invalid dates
            errors_found = False
            error_messages = []

            for index, value in df['DATE'].items():
                if pd.isnull(value):  # Check for empty cells
                    error_messages.append(f"TEST 9 DEFAULT INCORRECT: Empty cell found in row {index + 2}, column DATE.")
                    errors_found = True
                else:
                    date_str = str(value)
                    # Use regular expression to check for valid date format 'mm/dd/yyyy' or 'dd/mm/yyyy'
                    if not re.match(r'^(\d{1,2}/\d{1,2}/\d{4})|(\d{1,2}-\d{1,2}-\d{4})$', date_str):
                        error_messages.append(f"TEST 9 DEFAULT INCORRECT: Invalid date format found in row {index + 2}, column DATE: '{date_str}'.")
                        errors_found = True

            if not errors_found:
                print(f"TEST 9 DEFAULT CORRECT: Column Date format is correct in file '{file_path}'.")
            else:
                for message in error_messages:
                    print(message)
        except FileNotFoundError as fnf:
            print(f"Error: File '{file_path}' not found.")
        except Exception as e:
            print(f"Unexpected error while processing the file '{file_path}': {str(e)}")

# Call the function to perform the validation
if __name__ == '__main__':
    unittest.main()
