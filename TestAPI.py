import unittest
import pandas as pd
from datetime import datetime

class TestExcel(unittest.TestCase):

    def testCase1(self):

        #Description TestCase: Output Data Should remove shifts with Paycode LCUP, SHDIF and LUNCH
        #File: TMMC

        # Reade Excel file and select the sheet "OutputData"
        data_frame = pd.read_excel("TMMC W.E. 4.22.xlsx", sheet_name="OutputData")

        # Obtain the column 
        Column = data_frame["PAYCODE"]

        # Checked the expected Values
        expectedValues = ["LUNCH", "LCUP", "SHDIF"]

        missingValues = []

        #if an expected value is not found in column.values, it is added to the missing_values list.
        for value in expectedValues:
            if value in Column.values:
                missingValues.append(value)
        self.assertListEqual(missingValues, [], f"The following values are not present in the PAYCODE column in Excel TMMC file: {missingValues}")

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

    def testCase9(self):
        # Validate Output Out time from TMMC as it takes the out time
        # Archivo: TMMC
        data_frame = pd.read_excel("TMMC W.E. 4.22.xlsx", sheet_name="OutputData")

        nurse_column = data_frame["NAME"]
        startdtm_column = pd.to_datetime(data_frame["STARTDTM"])
        enddtm_column = pd.to_datetime(data_frame["ENDDTM"])

        nurse_name_1 = "Chekabab, Zahra"
        expected_value_enddtm_1 = datetime(2023, 4, 18, 7, 39)
        expected_value_startdtm_1 = datetime(2023, 4, 17, 18, 58)

        nurse_name_2 = "Maldonado, Marleny"
        expected_value_enddtm_2 = datetime(2023, 4, 20, 7, 26)
        expected_value_startdtm_2 = datetime(2023, 4, 19, 18, 55)

        errors = []

        found_nurse_1 = False
        found_nurse_2 = False

        for nurse, startdtm, enddtm in zip(nurse_column, startdtm_column, enddtm_column):
            if not found_nurse_1 and nurse == nurse_name_1:
                if enddtm != expected_value_enddtm_1:
                    errors.append(f"The value '{expected_value_enddtm_1}' is not present in ENDDTM for the nurse '{nurse_name_1}' in the TMM file.")
                if startdtm != expected_value_startdtm_1:
                    errors.append(f"The value '{expected_value_startdtm_1}' is not present in STARTDTM for the nurse '{nurse_name_1}' in the TMM file.")
                found_nurse_1 = True

            if not found_nurse_2 and nurse == nurse_name_2:
                if enddtm != expected_value_enddtm_2:
                    errors.append(f"The value '{expected_value_enddtm_2}' is not present in ENDDTM for the nurse '{nurse_name_2}' in the TMM file.")
                if startdtm != expected_value_startdtm_2:
                    errors.append(f"The value '{expected_value_startdtm_2}' is not present in STARTDTM for the nurse '{nurse_name_2}' in the TMM file.")
                found_nurse_2 = True

            if found_nurse_1 and found_nurse_2:
                break

        if errors:
            self.fail("\n".join(errors))

if __name__ == '__main__':
    unittest.main()
