import unittest
import pandas as pd

class TestExcel(unittest.TestCase):

    def testCase1(self):

        #Description TestCase: Output Data Should remove shifts with Paycode LCUP, SHDIF and LUNCH
        #File: TMMC

        # Lee el archivo de Excel y selecciona el sheet "OutputData"
        data_frame = pd.read_excel("TMMC W.E. 4.22.xlsx", sheet_name="OutputData")

        # Obtiene la columna deseada
        Column = data_frame["PAYCODE"]

        # Verifica los valores esperados
        expectedValues = ["LUNCH", "LCUP", "SHDIF"]

        missingValues = []

        #if an expected value is not found in column.values, it is added to the missing_values list.
        for value in expectedValues:
            if value in Column.values:
                missingValues.append(value)
        self.assertNotIn(missingValues, f"The following values are not present in the PAYCODE column in Excel TMMC file: {missingValues}")

    def testCase2(self):

        #Description TestCase: Remove SCHED shifts when is neccesary
        #File: DELTA

        data_frame = pd.read_excel("Delta Health 4.15.23.xlsx", sheet_name="OutputData")

        column = data_frame["PAYCODE"]

        expectedValues = ["SCHED"]

        for value in column:
            self.assertNotIn(value, expectedValues, f"The value '{value}' is present in the PAYCODE column in the DELTA file.")
    
    def testCase3(self):

        #Description TestCase: Remove SCHED shifts when is neccesary
        #File: TMMC
        data_frame = pd.read_excel("TMMC W.E. 4.22.xlsx", sheet_name="OutputData")

        Column = data_frame["PAYCODE"]

        expectedValues = ["SCHED"]

        for value in Column:
            self.assertNotIn(value, expectedValues, f"The value '{value}' is present in the PAYCODE column in the TMMC file.")

    
    def testCase7(self):
        #Validate Output has all nurses
        #File: DELTA modified
        data_frame = pd.read_excel("Delta Health 4.15.23.xlsx", sheet_name="OutputData")
        Column = data_frame["NAME"]
        expectedValues = ["Hunter, Angelique","Halums, Brittney","Cross, Destin","Radford, Gladys","Hale, Shannon","Kelly, Joby", "Lowe, Sherrie", "Lewis, Susan", "Towery, Brittany"]
        #Added a missing_values list to store the values that were not found in the "NAME" column.
        missingValues = []

        #if an expected value is not found in column.values, it is added to the missing_values list.
        for value in expectedValues:
            if value not in Column.values:
                missingValues.append(value)
        self.assertFalse(missingValues, f"The following values are not present in the NAME column in Excel Delta file: {missingValues}")

if __name__ == '__main__':
    unittest.main()
