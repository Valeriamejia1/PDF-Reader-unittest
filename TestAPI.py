import unittest
import pandas as pd

class TestExcel(unittest.TestCase):

    def testCase1(self):

        # Lee el archivo de Excel y selecciona el sheet "OutputData"
        data_frame = pd.read_excel("OutputAPI.xlsx", sheet_name="RawData")

        # Obtiene la columna deseada
        columna = data_frame["PAYCODE"]

        # Verifica los valores esperados
        valores_esperados = ["LUNVH", "LCUP", "SHDIF"]

        for valor in columna:
            self.assertNotIn(valor, valores_esperados, f"El valor '{valor}' está presente en la columna PAYCODE")

if __name__ == '__main__':
    unittest.main()
