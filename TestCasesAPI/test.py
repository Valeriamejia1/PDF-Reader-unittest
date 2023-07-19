import unittest
import pandas as pd

class TestExcelData(unittest.TestCase):
    def test_UKGC_4(self):
        excel_file = "TestCasesUKGCommon/Martin ppe 4.22.23.xlsx"
        sheet_name = "Sheet1"
        data = pd.read_excel(excel_file, sheet_name=sheet_name)
        name_column = "NAME"
        date_column = "DATE"
        Comment_column = "Comments"

        # Filtrar filas con "AUGUSTE, LOURDJINA" en la columna "NAME"
        filtered_rows = data[data[name_column] == "AUGUSTE, LOURDJINA"]

        for index, row in filtered_rows.iterrows():
            date_value = row[date_column]
            comment_value = row[Comment_column]

            expected_comment = "" if date_value == "04/22/2023" else "CC/MARTIN/MARTINNORTHHOSPITAL/CNO/NRS/INPAT-CLINICALDECISIONUNIT2W/RN"
            error_message = f"Error: The Comment value for row {index + 2} is not '{expected_comment}'"

            # Verificar si el valor de comentario es NaN
            if pd.isna(comment_value):
                comment_value = ""
                
            self.assertEqual(comment_value, expected_comment, error_message)

    def test_UKGC_5(self):
        excel_file = "TestCasesUKGCommon/Martin ppe 4.22.23.xlsx"
        sheet_name = "Sheet1"
        data = pd.read_excel(excel_file, sheet_name=sheet_name)
        name_column = "NAME"
        date_column = "DATE"
        glcode_column = "GLCODE"

        # Filtrar filas con "AUGUSTE, LOURDJINA" en la columna "NAME"
        filtered_rows = data[data[name_column] == "AUGUSTE, LOURDJINA"]

        for index, row in filtered_rows.iterrows():
            date_value = row[date_column]
            glcode_value = row[glcode_column]

            expected_glcode = "3100-3100-31812" if date_value == "04/22/2023" else "3100-3100-31429"
            error_message = f"Error: The GLCODE value for row {index + 2} is not '{expected_glcode}'"
            self.assertEqual(glcode_value, expected_glcode, error_message)
if __name__ == '__main__':
    unittest.main()

