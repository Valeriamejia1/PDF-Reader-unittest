import unittest
import pandas as pd
import string
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.filterwarnings("ignore", category=DeprecationWarning)
import numpy as np

class TestExcel(unittest.TestCase):

    def test_UKGK_1(self):

        # Loads the Excel file into a DataFram
        df = pd.read_excel('OUTPUT UKGKronos/UKG Kronos empty.xlsx', header=None)

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
            excel_file = 'OUTPUT UKGKronos/Qualivis Time report PPE 062423.xlsx'
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

    #Methods required for test_UKGK4

    def test_UKGS_4(self):
        self.compare_excel_files("TestCasesUKGKronos/Qualivis Time report PPE 062423 ORIG.xlsx", "OUTPUT UKGKronos/Qualivis Time report PPE 062423.xlsx")
        print("TEST 4 UKGKronos CORRECT: Qualivis Time report PPE 062423.xlsx data match the original version")

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

    #Descrition: Check Exe has the same data as last commit
    #Files: Qualivis Time report PPE 062423

    #Methods required for test_UKGK4

    def verified_columns(path_file, sheet_namee, column):
        try:
            # Leer el archivo Excel en un DataFrame
            df = pd.read_excel(path_file, sheet_name=sheet_namee, engine='openpyxl')

            # Verificar si la columna está presente
            assert column in df.columns, f'The column "{column}" is not present in the file.'

            # Verificar si hay filas con campos vacíos en la columna
            filas_vacias = df[df[column].isnull()]
            if not filas_vacias.empty:
                mensaje_error = 'There are empty fields in the column "{0}" in the following rows:\n'.format(column)
                for index, _ in filas_vacias.iterrows():
                    mensaje_error += 'Row(s): {0}\n'.format(index + 2) 
                # Obtener la letra de la columna correspondiente
                col_idx = df.columns.get_loc(column)
                col_letra = string.ascii_uppercase[col_idx]
                mensaje_error += 'Column: {0}\n'.format(col_letra)
                raise AssertionError(mensaje_error)

            return True
        except Exception as e:
            print(f'Error: {e}')
            return False

    def test_UKGK_5(self):
        path_file = 'OUTPUT UKGKronos/Qualivis Time report PPE 062423.xlsx'
        sheet_namee = 'Sheet1'
        column = 'PRIMARY JOB'

        resultado = TestExcel.verified_columns(path_file, sheet_namee, column)
        print("TEST 5 UKGKronos CORRECT: Qualivis Time report PPE 062423 CORRECT: Qualivis Time report PPE 062423.xlsx column PRIMARY JOB  is in the file")
        self.assertTrue(resultado)
    
    def test_UKGK_6(self):
        # Load the Excel files into pandas DataFrames
        expected_df = pd.read_excel("TestCasesUKGKronos/Qualivis Time report PPE 062423 ORIG.xlsx")
        actual_df = pd.read_excel("OUTPUT UKGKronos/Qualivis Time report PPE 062423.xlsx")

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

    def verified_names(self, pahtfile):
        try:
            # Leer el archivo Excel en un DataFrame sin considerar el encabezado
            df = pd.read_excel(pahtfile, sheet_name='Sheet1', header=None, engine='openpyxl')

            # Nombre esperado y rangos de filas para cada nombre
            expectedname = {
                'VALLE, DIANA DEJESUS': (1026, 1068),
                'JISON, ROSE ANNE': (413, 420),
                'BURNS-BARRINO,LATASHA': (105, 111)
            }

            # Verificar cada nombre en su rango de filas especificado
            for name, (fila_inicial, fila_final) in expectedname.items():
                offset = 1 if df.iat[0, 1] == 'NAME' else 0  # Ajuste para considerar si hay encabezado
                for index in range(fila_inicial - offset, fila_final - 1):
                    if df.at[index, 1] != name:  # Columna 'NAME' es la columna 1
                        raise AssertionError(f"Error: The name '{name}' in row {index + 1} is different. Value found: {df.at[index, 1]}")

            return True

        except Exception as e:
            print(f'Error: {e}')
            return False

    def test_UKGK_7(self):
        pahtfile = 'OUTPUT UKGKronos/Qualivis Time report PPE 062423.xlsx'

        resultado = self.verified_names(pahtfile)
        if resultado:
            print("TEST 7 UKGKronos CORRECT: The names 'VALLE, DIANA DEJESUS', 'JISON, ROSE ANNE' and 'BURNS-BARRINO, LATASHA' are present and spelled correctly in the specified rows.")
        self.assertTrue(resultado)

    def test_UKGK_8(self):
        # Load the Excel file into a pandas DataFrame
        df = pd.read_excel("OUTPUT UKGKronos/Qualivis Time report PPE 062423.xlsx")

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
        df = pd.read_excel("OUTPUT UKGKronos/Qualivis Time report PPE 062423.xlsx")

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

    def test_UKGK10(self):
        # Define the expected values
        expected_name = "DUNDAS, TAYLOR"
        expected_date = "06/17/2023"
        expected_paycode = "OR-HRT CALL WE"

        # Load the Excel file into a pandas DataFrame
        file_path = "OUTPUT UKGKronos/Qualivis Time report PPE 062423.xlsx"
        sheet_name = "Sheet1"
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Function to check and report discrepancies in the DataFrame
        errors = []
        for index, row in df.iterrows():
            if row["NAME"] == expected_name and row["DATE"] == expected_date:
                actual_paycode = str(row["PAYCODE"]).strip()  # Remove spaces
                if actual_paycode == expected_paycode:
                    # Found a matching row, exit the loop since we found what we were looking for
                    break

                else:
                    errors.append(f"Row {index} has PAYCODE mismatch. Expected: {expected_paycode}, Actual: {actual_paycode}")
                    break  # Added a break to exit the loop if there's a mismatch

        # If there are any errors in the DataFrame, fail the test
        if errors:
            self.fail("\n".join(errors))

        # If no errors were found, the test was successful
        print("TEST 10 UKGKronos CORRECT: The PAYCODE column is correct for the nurse that you requested.")


if __name__ == "__main__":
    unittest.main()






