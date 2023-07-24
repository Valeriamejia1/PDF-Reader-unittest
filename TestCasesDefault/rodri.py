import unittest
import pandas as pd
import re

class ExcelTestCase(unittest.TestCase):

    def test_default_11(self): 
        # Files to validate
        #Desc: Date with Dash needs to be display
        file_paths = ['TestCasesDefault\\Scripps Approved Kronos we 6.25.22 minutes.xlsx', 'TestCasesDefault\\Scripps Approved Kronos we 6.25.22.xlsx']
        sheet_name = 'Sheet1'

        # Regular expression to check for valid date format 'mm/dd/yyyy' and to detect '+' or '-' symbols
        date_pattern = r'^\d{1,2}/\d{1,2}/\d{4}(?<![-+])$'

        all_files_good = True  # Initialize to True
        error_messages = []

        for file_path in file_paths:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)

                # Check for missing or invalid dates
                for index, value in df['DATE'].items():
                    if pd.isnull(value):  # Check for empty cells
                        error_messages.append(f"TEST 11 DEFAULT INCORRECT: Empty cell found in row {index + 2}, column DATE in file '{file_path}'.")
                        all_files_good = False  # Update to False if any file is incorrect
                    else:
                        date_str = str(value)
                        if not re.match(date_pattern, date_str):
                            error_messages.append(f"TEST 11 DEFAULT INCORRECT: Invalid date format found in row {index + 2}, column DATE: '{date_str}' in file '{file_path}'.")
                            all_files_good = False  # Update to False if any file is incorrect

            except FileNotFoundError as fnf:
                error_messages.append(f"Error: File '{file_path}' not found.")
                all_files_good = False
            except Exception as e:
                error_messages.append(f"Unexpected error while processing the file '{file_path}': {str(e)}")
                all_files_good = False

        for message in error_messages:
            print(message)

        if all_files_good:
            print(f"TEST 11 DEFAULT CORRECT: Column Date format is correct in all files.")

        if not all_files_good:
            self.fail()

# Call the function to perform the validation
if __name__ == '__main__':
    unittest.main()
