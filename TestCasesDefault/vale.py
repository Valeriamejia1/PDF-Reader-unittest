import unittest
import pandas as pd
import re

class TestGLCode(unittest.TestCase):
    def test_DEFAULT_12(self):
        # Add the 'self' parameter
        # List of file paths and sheet names to validate
        file_data = [
            {'file_path': 'TestCasesDefault/Time Detail_100822-102122 minutes.xlsx', 'sheet_name': 'Sheet1'},
            {'file_path': 'TestCasesDefault/Time Detail_100822-102122.xlsx', 'sheet_name': 'Sheet1'}
        ]

        # List to store all error messages
        all_error_messages = []
        correct_files = []

        for data in file_data:
            file_path = data['file_path']
            sheet_name = data['sheet_name']

            # Read the Excel file and the specified sheet
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)

                # Check for missing or invalid dates
                error_messages = []  # Collect error messages for this file
                for index, value in df['DATE'].items():
                    if pd.isnull(value):  # Check for empty cells
                        error_messages.append(f"TEST 12 DEFAULT INCORRECT: Empty cell found in row {index + 2}, column DATE. File: {file_path}")
                    else:
                        date_str = str(value)
                        # Use regular expression to check for valid date format 'mm/dd/yyyy' or 'dd/mm/yyyy'
                        if not re.match(r'^(\d{1,2}/\d{1,2}/\d{4})|(\d{1,2}-\d{1,2}-\d{4})$', date_str):
                            error_messages.append(f"TEST 12 DEFAULT INCORRECT: Invalid date format found in row {index + 2}, column DATE: '{date_str}'. File: {file_path}")

                if error_messages:  # If there are errors for this file, add them to the all_error_messages list
                    all_error_messages.extend(error_messages)
                else:
                    correct_files.append(file_path)
            except FileNotFoundError as fnf:
                self.fail(f"Error: File '{file_path}' not found.")
            except Exception as e:
                self.fail(f"Unexpected error while processing the file '{file_path}': {str(e)}")

        # Raise a single failure with all the collected error messages
        if all_error_messages:
            error_message = "\n".join(all_error_messages)
            self.fail(error_message)
        else:
            if correct_files:
                correct_file_message = "\n".join(correct_files)
                print(f"TEST 12 DEFAULT CORRECT: Column Date format is correct in both files")

if __name__ == "__main__":
    unittest.main()
