import unittest
import pandas as pd

class TestExcel(unittest.TestCase):

    def testCase1(self):

        #Description TestCase: Output Data Should remove shifts with Paycode LCUP, SHDIF and LUNCH
        #File: TMMC

        # Lee el archivo de Excel y selecciona el sheet "OutputData"
        data_frame = pd.read_excel("TMMC W.E. 4.22.xlsx", sheet_name="OutputData")

        # Obtiene la columna deseada
        columna = data_frame["PAYCODE"]

        # Verifica los valores esperados
        valores_esperados = ["LUNCH", "LCUP", "SHDIF"]

        for valor in columna:
            self.assertNotIn(valor, valores_esperados, f"El valor '{valor}' está presente en la columna PAYCODE en el File TMMC")

    def testCase2(self):

        #Description TestCase: Remove SCHED shifts when is neccesary
        #File: DELTA

        # Lee el archivo de Excel y selecciona el sheet "OutputData"
        data_frame = pd.read_excel("Delta Health 4.15.23.xlsx", sheet_name="OutputData")

        # Obtiene la columna deseada
        columna = data_frame["PAYCODE"]

        # Verifica los valores esperados
        valores_esperados = ["SCHED"]

        for valor in columna:
            self.assertNotIn(valor, valores_esperados, f"El valor '{valor}' está presente en la columna PAYCODE en el File DELTA")
    
    def testCase3(self):

        #Description TestCase: Remove SCHED shifts when is neccesary
        #File: TMMC

        # Lee el archivo de Excel y selecciona el sheet "OutputData"
        data_frame = pd.read_excel("TMMC W.E. 4.22.xlsx", sheet_name="OutputData")

        # Obtiene la columna deseada
        columna = data_frame["PAYCODE"]

        # Verifica los valores esperados
        valores_esperados = ["SCHED"]

        for valor in columna:
            self.assertNotIn(valor, valores_esperados, f"El valor '{valor}' está presente en la columna PAYCODE en el File TMMC")


if __name__ == '__main__':
    unittest.main()
