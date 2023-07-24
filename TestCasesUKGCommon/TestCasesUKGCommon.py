import unittest
import pandas as pd

class ExcelTest(unittest.TestCase):

    def test_UKGC_1(self):

        # Loads the Excel file into a DataFram
        df = pd.read_excel('TestCasesUKGCommon/UKG Common Empty.xlsx', header=None)

        # Gets the number of rows with data beyond the headers
        num_data_rows = len(df) - 1 

        # Check if there are additional rows with data
        if num_data_rows > 0:

            self.fail("There are {} additional rows with data in the Excel file".format(num_data_rows))
        
        print(".TEST 1 UKGCommon CORRECT: The UKG Common Empty.xlsx file not contains additional rows to the header")
    
    def test_UKGC_2(self):
        # Description: Check last shift is present
        # File: Martin ppe
        # Upload Excel file
        excel_file = 'TestCasesUKGCommon/Martin ppe 4.22.23.xlsx'
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
                print("TEST 2 UKGCommon CORRECT: Data for WARREN, DANIELLE was found and is correct in file Martin ppe 4.22.23.")
   
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
    #Files: Martin ppe 4.22.23, martin b, time martin a ORIG

    def test_UKGC_3_1(self):
        self.compare_excel_files("TestCasesUKGCommon/Martin ppe 4.22.23 ORIG.xlsx", "TestCasesUKGCommon/Martin ppe 4.22.23.xlsx")
        print("TEST 3.1 UKGCommon CORRECT: Martin ppe 4.22.23.xlsx data match the original version")

    def test_UKGC_3_2(self):
        self.compare_excel_files("TestCasesUKGCommon/martin b ORIG.xlsx", "TestCasesUKGCommon/martin b.xlsx")
        print("TEST 3.2 UKGCommon CORRECT: martin b ORIG.xlsx data match the original version")

    def test_UKGC_3_3(self):
        self.compare_excel_files("TestCasesUKGCommon/martin time a ORIG.xlsx", "TestCasesUKGCommon/martin time a.xlsx")
        print("TEST 3.3 UKGCommon CORRECT: martin time a ORIG.xlsx data match the original version")

    def test_UKGC_4(self):
        excel_file = "TestCasesUKGCommon/Martin ppe 4.22.23.xlsx"
        sheet_name = "Sheet1"
        data = pd.read_excel(excel_file, sheet_name=sheet_name)
        name_column = "NAME"
        date_column = "DATE"
        Comment_column = "Comments"

        # Filter rows with "AUGUSTE, LOURDJINA" in column "NAME".
        filtered_rows = data[data[name_column] == "AUGUSTE, LOURDJINA"]

        for index, row in filtered_rows.iterrows():
            date_value = row[date_column]
            comment_value = row[Comment_column]

            expected_comment = "" if date_value == "04/22/2023" else "CC/MARTIN/MARTINNORTHHOSPITAL/CNO/NRS/INPAT-CLINICALDECISIONUNIT2W/RN"
            error_message = f"Error: The Comment value for row {index + 2} is not '{expected_comment}'"

            # Check if the comment value is NaN
            if pd.isna(comment_value):
                comment_value = ""
                
            self.assertEqual(comment_value, expected_comment, error_message)

        print("TEST 4 UKGCommon CORRECT: For AUGUSTE, LOURDJINA the comments match those in the PDF")

    def test_UKGC_5(self):
        excel_file = "TestCasesUKGCommon/Martin ppe 4.22.23.xlsx"
        sheet_name = "Sheet1"
        data = pd.read_excel(excel_file, sheet_name=sheet_name)
        name_column = "NAME"
        date_column = "DATE"
        glcode_column = "GLCODE"

        # Filter rows with "AUGUSTE, LOURDJINA" in column "NAME".
        filtered_rows = data[data[name_column] == "AUGUSTE, LOURDJINA"]

        for index, row in filtered_rows.iterrows():
            date_value = row[date_column]
            glcode_value = row[glcode_column]

            expected_glcode = "3100-3100-31812" if date_value == "04/22/2023" else "3100-3100-31429"
            error_message = f"Error: The GLCODE value for row {index + 2} is not '{expected_glcode}'"
            self.assertEqual(glcode_value, expected_glcode, error_message)

        print ("TEST 5 UKGCommon CORRECT: For AUGUSTE, LOURDJINA the GLCODES match those in the PDF")

if __name__ == '__main__':
    unittest.main()
