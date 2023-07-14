import unittest
import pandas as pd

class ExcelDataTest(unittest.TestCase):
    def test_API_5(self):
        # Cargar el archivo de Excel
        file_path = 'TestCases API\API Empty.xlsx'
        sheet_name = 'OutputData'

        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Verificar si hay datos en alguna fila después de la primera
        if len(df) > 1:
            self.fail("The API Empty.xlsx file contains additional rows to the header")

        # Si no se encontraron datos, entonces el archivo solo tiene el encabezado
        self.assertTrue(len(df) == 1, "The API Empty.xlsx file contains data")

if __name__ == '__main__':
    unittest.main()
