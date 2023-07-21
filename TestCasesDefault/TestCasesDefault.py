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
    
    def validar_glcode(self, glcode):
        # Eliminar los guiones "-" y contar los dígitos restantes
        glcode_sin_guiones = glcode.replace('-', '')
        return len(glcode_sin_guiones) == 9 and glcode_sin_guiones.isdigit()

    def test_DEFAULT_3(self):
        archivos = ["TestCasesDefault/Combined File minutes.xlsx", "TestCasesDefault/Combined File.xlsx"]

        all_incorrect_rows = set()  # Usamos un conjunto para almacenar las filas incorrectas sin duplicados

        for archivo in archivos:
            try:
                with pd.ExcelFile(archivo) as xls:
                    sheet_name = xls.sheet_names[0]  # Asumimos que la hoja de interés es la primera

                    # Cargar el archivo y convertir la columna "GLCODE" al formato "text"
                    df = pd.read_excel(xls, sheet_name)
                    df["GLCODE"] = df["GLCODE"].astype(str)

                    # Validar que todas las celdas en la columna "GLCODE" contengan 9 dígitos sin guiones "-"
                    incorrect_rows = []
                    for index, glcode in df["GLCODE"].items():
                        if not self.validar_glcode(glcode):
                            incorrect_rows.append((archivo, index + 2, glcode))

                    # Agregar las filas incorrectas al conjunto general
                    all_incorrect_rows.update(incorrect_rows)
            except Exception as e:
                # Si ocurre una excepción, muestra el mensaje de error y registra la excepción.
                print(f"Error al procesar el archivo {archivo}: {str(e)}")

        # Mostrar mensaje de fallo con todas las filas incorrectas de todos los archivos
        if all_incorrect_rows:
            error_msg = "El GLCODE no contiene 9 digitos en las siguientes filas y archivos:\n"
            for archivo, fila, glcode in all_incorrect_rows:
                error_msg += f"Archivo: {archivo}, Fila: {fila}, GLCODE: {glcode}\n"
            self.fail(error_msg)

        print(".TEST 3 DEFAULT CORRECT: All GLCODEs have 9 numeric digits.")




    
    def test_Default_4_1(self):

        files = ["TestCasesDefault/6-11.xlsx", "TestCasesDefault/6-11 Minutes.xlsx"]

        all_errors = []  # Lista para almacenar todos los errores encontrados

        for archivo in files:
            xls = pd.ExcelFile(archivo)
            sheet_name = xls.sheet_names[0]  # Asumimos que la hoja de interés es la primera

            # Cargar el archivo y filtrar por "Attaway, Brooke" en la columna "NAME"
            df = pd.read_excel(xls, sheet_name)
            brooke_df = df[df["NAME"] == "Attaway, Brooke"]

            # Verificar que en la columna "GLCODE" esté alguno de los valores esperados
            incorrect_rows = []
            for index, row in brooke_df.iterrows():
                glcode = str(row["GLCODE"])
                if glcode not in ["6142", "006142"]:
                    incorrect_rows.append((archivo, index + 2, glcode))

            # Agregar los errores al listado general
            all_errors.extend(incorrect_rows)

            xls.close()  # Cerrar el archivo después de leerlo

        # Mostrar mensaje de fallo con todos los errores encontrados
        if all_errors:
            error_msg = "Errores para Attaway, Brooke:\n"
            for archivo, fila, glcode in all_errors:
                error_msg += f"Archivo: {archivo}, Fila: {fila}, GLCODE: {glcode}\n"
            self.fail(error_msg)

        print("TEST 4.1 DEFAULT CORRECT: FILE TEST 6-11 BROOKE CORRECT: All GLCODEs for Brooke are valid.")

    def test_Default_4_2(self):    
        files = ["TestCasesDefault/6-11.xlsx", "TestCasesDefault/6-11 Minutes.xlsx"]

        all_errors = []  # Lista para almacenar todos los errores encontrados

        for archivo in files:
            xls = pd.ExcelFile(archivo)
            sheet_name = xls.sheet_names[0]  # Asumimos que la hoja de interés es la primera

            # Cargar el archivo y filtrar por "Barr, Brieann" en la columna "NAME"
            df = pd.read_excel(xls, sheet_name)
            brieann_df = df[df["NAME"] == "Barr, Brieann"]

            # Verificar que en la columna "GLCODE" esté alguno de los valores esperados
            incorrect_rows = []
            for index, row in brieann_df.iterrows():
                glcode = str(row["GLCODE"])
                if glcode not in ["007317", "6402", "7317"]:
                    incorrect_rows.append((archivo, index + 2, glcode))

            # Agregar los errores al listado general
            all_errors.extend(incorrect_rows)

            xls.close()  # Cerrar el archivo después de leerlo

        # Mostrar mensaje de fallo con todos los errores encontrados
        if all_errors:
            error_msg = "Errores para Barr, Brieann:\n"
            for archivo, fila, glcode in all_errors:
                error_msg += f"Archivo: {archivo}, Fila: {fila}, GLCODE: {glcode}\n"
            self.fail(error_msg)

        print("TEST 4.2 DEFAULT CORRECT: FILE TEST 6-11 BRIEANN CORRECT: All GLCODEs for Brieann are valid.")

    def test_DEFAULT_5_1(self):
    # Excel files you wish to validate along with their names
        filenames = [
            ("TestCasesDefault/1667243700213_980358248.xlsx", "1667243700213_980358248.xlsx"),
            ("TestCasesDefault/1667243700213_980358248 minutes.xlsx", "1667243700213_980358248 minutes.xlsx")
        ]

        # Values you want to search for in each file
        name_value = "Tr-Freeman, Alexander"
        glcode_value = "1110.2115.1057"

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
        
        print(".TEST 5.1 DEFAULT CORRECT: Alexander's GLCODE contains . and is 1110.2115.1057")

    def test_DEFAULT_5_2(self):
        # Excel files you wish to validate along with their names
            filenames = [
                ("TestCasesDefault/1667243700213_980358248.xlsx", "1667243700213_980358248.xlsx"),
                ("TestCasesDefault/1667243700213_980358248 minutes.xlsx", "1667243700213_980358248 minutes.xlsx")
            ]

            # Values you want to search for in each file
            name_value = "Tr-Belcher, Hanna"
            glcode_value = "1130.2305.1474"

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
            
            print(".TEST 5.2 DEFAULT CORRECT: Hannas's GLCODE contains . and is 1110.2115.1057")

    def test_default_11(self):
        # File path and name of the Excel file
        excel_file = 'TestCasesDefault/6-11.xlsx'
        
        # Sheet name in the Excel file
        excel_sheet = 'Sheet1'
        
        # Names to search in the NAME column
        names_to_search = ['Aguilar, Isabelle', 'Aguilar, Marissa']
        
        # Columns to check for empty values
        columns_to_check = ['DATE', 'GLCODE', 'PAYCODE', 'STARTDTM', 'ENDDTM', 'HOURS']
        
        # Read the Excel file into a DataFrame
        df = pd.read_excel(excel_file, sheet_name=excel_sheet)
        
        # Filter by names in the NAME column
        filtered_df = df[df['NAME'].isin(names_to_search)]
        
        # Check for empty columns in the specified columns
        empty_columns = filtered_df[filtered_df[columns_to_check].isnull().any(axis=1)]
        
        # Generate the error message
        error_msg = "The following rows and columns have empty data:\n"
        for idx, row in empty_columns.iterrows():
            for col in columns_to_check:
                if pd.isnull(row[col]):
                    error_msg += f"Column: {col}, Row: {idx+2}\n"
        
        # Assert that there are no empty columns
        self.assertTrue(empty_columns.empty, error_msg)
        print(".TEST 11 DEFAULT CORRECT: All the information of the nurses are in the file.")
    
    def test_Default_12(self):
        # File path and name of the Excel file
        excel_file = 'TestCasesDefault/6-11.xlsx'
        
        # Sheet name in the Excel file
        excel_sheet = 'Sheet1'
        
        # Name to search in the NAME column
        name_to_search = 'Anderson, Jennifer'
        
        # Read the Excel file into a DataFrame
        df = pd.read_excel(excel_file, sheet_name=excel_sheet)
        
        # Count the occurrences of the name in the NAME column
        name_count = df['NAME'].value_counts().get(name_to_search, 0)
        
        # Check if the name appears exactly three times
        self.assertEqual(name_count, 3, f"The name '{name_to_search}' registered {name_count} shifts, but it should registered 3 shifts.")
        
        # Print a message if the test passes
        print(".TEST 12 DEFAULT CORRECT: The nurse registered in the file three shifts.")

if __name__ == '__main__':

    unittest.main()
