import unittest
import pandas as pd
import numpy as np

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
        data_frame = pd.read_excel("Delta Health 4.15.23 SCHED.xlsx", sheet_name="OutputData")
        Column = data_frame["NAME"]
        expectedValues = ["Hunter, Angelique","Halums, Brittney","Cross, Destin","Radford, Gladys","Hale, Shannon","Kelly, Joby", "Lowe, Sherrie", "Lewis, Susan", "Towery, Brittany"]
        #Added a missing_values list to store the values that were not found in the "NAME" column.
        missingValues = []

        #if an expected value is not found in column.values, it is added to the missing_values list.
        for value in expectedValues:
            if value not in Column.values:
                missingValues.append(value)
        self.assertFalse(missingValues, f"The following values are not present in the NAME column in Excel Delta file: {missingValues}")

    def testcase9(self):
        #Validate Output Out time from TMMC as it takes the out time
        #File: TMMC WITH SCHED

        # Upload Excel file
        df = pd.read_excel("TMMC W.E. 4.22 SCHED.xlsx", sheet_name="OutputData")

        # Format the date and time columns to match the expected format.
        df["STARTDTM"] = pd.to_datetime(df["STARTDTM"], format="%m/%d/%Y %H:%M")
        df["ENDDTM"] = pd.to_datetime(df["ENDDTM"], format="%m/%d/%Y %H:%M")
        
        # Validate data for the nurse "Chekabab, Zahra"
        nurse_name = "Chekabab, Zahra"
        
        nurse_data = df[df["NAME"] == nurse_name]
        expected_data = [
            ("REG", pd.to_datetime("2023-04-17 18:58:00"), pd.to_datetime("2023-04-18 07:39:00")),
            ("SCHED", pd.to_datetime("2023-04-18 19:01:00"), pd.to_datetime("2023-04-19 07:34:00"))
        ]
        
        for expected_paycode, expected_startdtm, expected_enddtm in expected_data:
            self.assertTrue(any((nurse_data["PAYCODE"] == expected_paycode) & (nurse_data["STARTDTM"] == expected_startdtm) & (nurse_data["ENDDTM"] == expected_enddtm)), f"No se encontró la línea con los valores esperados para {nurse_name} - PAYCODE: {expected_paycode}, STARTDTM: {expected_startdtm}, ENDDTM: {expected_enddtm}")
        
        # Validate data for the nurse "Maldonado, Marleny".
        nurse_name = "Maldonado, Marleny"
        
        nurse_data = df[df["NAME"] == nurse_name]
        expected_data = [
            ("REG", pd.to_datetime("2023-04-19 18:55:00"), pd.to_datetime("2023-04-20 07:26:00")),
            ("SCHED", pd.to_datetime("2023-04-24 19:00:00"), np.nan)
        ]
        
        for expected_paycode, expected_startdtm, expected_enddtm in expected_data:
            if pd.isnull(expected_enddtm):
                self.assertTrue(any((nurse_data["PAYCODE"] == expected_paycode) & (nurse_data["STARTDTM"] == expected_startdtm) & pd.isnull(nurse_data["ENDDTM"])), f"The line with the expected values was not found for {nurse_name} - PAYCODE: {expected_paycode}, STARTDTM: {expected_startdtm}, ENDDTM: NaN")
            else:
                self.assertTrue(any((nurse_data["PAYCODE"] == expected_paycode) & (nurse_data["STARTDTM"] == expected_startdtm) & (nurse_data["ENDDTM"] == expected_enddtm)), f"The line with the expected values was not found for {nurse_name} - PAYCODE: {expected_paycode}, STARTDTM: {expected_startdtm}, ENDDTM: {expected_enddtm}")



if __name__ == '__main__':
    unittest.main()
