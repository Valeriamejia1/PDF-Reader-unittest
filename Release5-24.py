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

        #Descrition: Validate Output has all nurses
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
        # Description: Validate Output Out time from TMMC as it takes the out time
        # File: TMMC WITH SCHED

        # Upload Excel file
        df = pd.read_excel("TMMC W.E. 4.22 SCHED.xlsx", sheet_name="OutputData")

        # Format the date and time columns to match the expected format.
        df["STARTDTM"] = pd.to_datetime(df["STARTDTM"], format="%m/%d/%Y %H:%M")
        df["ENDDTM"] = pd.to_datetime(df["ENDDTM"], format="%m/%d/%Y %H:%M")

        errors = []  # List for storing error messages

        # Validate data for the nurse "Chekabab, Zahra"
        nurse_name = "Chekabab, Zahra"

        nurse_data = df[df["NAME"] == nurse_name]
        expected_data = [
            ("REG", pd.to_datetime("2023-04-17 18:58:00"), pd.to_datetime("2023-04-18 07:39:00")),
            ("SCHED", pd.to_datetime("2023-04-18 19:01:00"), pd.to_datetime("2023-04-19 07:34:00"))
        ]

        for expected_paycode, expected_startdtm, expected_enddtm in expected_data:
            if not any((nurse_data["PAYCODE"] == expected_paycode) & (nurse_data["STARTDTM"] == expected_startdtm) & (nurse_data["ENDDTM"] == expected_enddtm)):
                error_message = f"The line with the expected values was not found for {nurse_name} - PAYCODE: {expected_paycode}, STARTDTM: {expected_startdtm}, ENDDTM: {expected_enddtm}"
                errors.append(error_message)

        # Validate data for the nurse "Maldonado, Marleny".
        nurse_name = "Maldonado, Marleny"

        nurse_data = df[df["NAME"] == nurse_name]
        expected_data = [
            ("REG", pd.to_datetime("2023-04-19 18:55:00"), pd.to_datetime("2023-04-20 07:26:00")),
            ("SCHED", pd.to_datetime("2023-04-24 19:00:00"), np.nan)
        ]

        for expected_paycode, expected_startdtm, expected_enddtm in expected_data:
            if pd.isnull(expected_enddtm):
                if not any((nurse_data["PAYCODE"] == expected_paycode) & (nurse_data["STARTDTM"] == expected_startdtm) & pd.isnull(nurse_data["ENDDTM"])):
                    error_message = f"The line with the expected values was not found for {nurse_name} - PAYCODE: {expected_paycode}, STARTDTM: {expected_startdtm}, ENDDTM: NaN"
                    errors.append(error_message)
            else:
                if not any((nurse_data["PAYCODE"] == expected_paycode) & (nurse_data["STARTDTM"] == expected_startdtm) & (nurse_data["ENDDTM"] == expected_enddtm)):
                    error_message = f"The line with the expected values was not found for {nurse_name} - PAYCODE: {expected_paycode}, STARTDTM: {expected_startdtm}, ENDDTM: {expected_enddtm}"
                    errors.append(error_message)

        # Check for cumulative errors and display them as a single assertion at the end
        self.assertEqual(errors, [], f"The following errors were found in the testcase:\n\n{', '.join(errors)}")

    def testcase10(self):
        # Descrition: Check RAW data Out has the same hour as output
        # File: TMMC WITH SCHED

        # Upload Excel file
        df_output = pd.read_excel("TMMC W.E. 4.22 SCHED.xlsx", sheet_name="OutputData")
        df_raw = pd.read_excel("TMMC W.E. 4.22 SCHED.xlsx", sheet_name="RawData")

        # Format the date and time columns to match the expected format.
        df_output["STARTDTM"] = pd.to_datetime(df_output["STARTDTM"], format="%m/%d/%Y %H:%M")
        df_output["ENDDTM"] = pd.to_datetime(df_output["ENDDTM"], format="%m/%d/%Y %H:%M")
        df_raw["STARTDTM"] = pd.to_datetime(df_raw["STARTDTM"], format="%m/%d/%Y %H:%M")

        errors = []  # List for storing error messages

        # Validate data for the nurse "Chekabab, Zahra"
        nurse_name = "Chekabab, Zahra"

        nurse_data_output = df_output[df_output["NAME"] == nurse_name]
        nurse_data_raw = df_raw[df_raw["NAME"] == nurse_name]
        expected_data_output = [
            ("REG", pd.to_datetime("2023-04-17 18:58:00"), pd.to_datetime("2023-04-18 07:39:00")),
            ("SCHED", pd.to_datetime("2023-04-18 19:01:00"), pd.to_datetime("2023-04-19 07:34:00"))
        ]
        expected_data_raw = [
            ("REG", pd.to_datetime("2023-04-17 18:58:00")),
            ("SCHED", pd.to_datetime("2023-04-18 19:01:00"))
        ]

        for expected_paycode, expected_startdtm, expected_enddtm in expected_data_output:
            if not any((nurse_data_output["PAYCODE"] == expected_paycode) & (nurse_data_output["STARTDTM"] == expected_startdtm) & (nurse_data_output["ENDDTM"] == expected_enddtm)):
                error_message = f"The line with the expected values was not found for {nurse_name} in 'OutputData' - PAYCODE: {expected_paycode}, STARTDTM: {expected_startdtm}, ENDDTM: {expected_enddtm}"
                errors.append(error_message)

        for expected_paycode, expected_startdtm in expected_data_raw:
            if not any((nurse_data_raw["PAYCODE"] == expected_paycode) & (nurse_data_raw["STARTDTM"] == expected_startdtm)):
                error_message = f"The line with the expected values was not found for {nurse_name} in 'RawData' - PAYCODE: {expected_paycode}, STARTDTM: {expected_startdtm}"
                errors.append(error_message)

        # Validate data for the nurse "Maldonado, Marleny".
        nurse_name = "Maldonado, Marleny"

        nurse_data_output = df_output[df_output["NAME"] == nurse_name]
        nurse_data_raw = df_raw[df_raw["NAME"] == nurse_name]
        expected_data_output = [
            ("REG", pd.to_datetime("2023-04-19 18:55:00"), pd.to_datetime("2023-04-20 07:26:00")),
            ("SCHED", pd.to_datetime("2023-04-24 19:00:00"), pd.NaT)
        ]
        expected_data_raw = [
            ("REG", pd.to_datetime("2023-04-19 18:55:00"))
        ]

        for expected_paycode, expected_startdtm, expected_enddtm in expected_data_output:
            if pd.isnull(expected_enddtm):
                if not any((nurse_data_output["PAYCODE"] == expected_paycode) & (nurse_data_output["STARTDTM"] == expected_startdtm) & pd.isnull(nurse_data_output["ENDDTM"])):
                    error_message = f"The line with the expected values was not found for {nurse_name} in 'OutputData' - PAYCODE: {expected_paycode}, STARTDTM: {expected_startdtm}, ENDDTM: NaN"
                    errors.append(error_message)
            else:
                if not any((nurse_data_output["PAYCODE"] == expected_paycode) & (nurse_data_output["STARTDTM"] == expected_startdtm) & (nurse_data_output["ENDDTM"] == expected_enddtm)):
                    error_message = f"The line with the expected values was not found for {nurse_name} in 'OutputData' - PAYCODE: {expected_paycode}, STARTDTM: {expected_startdtm}, ENDDTM: {expected_enddtm}"
                    errors.append(error_message)

        for expected_paycode, expected_startdtm in expected_data_raw:
            if not any((nurse_data_raw["PAYCODE"] == expected_paycode) & (nurse_data_raw["STARTDTM"] == expected_startdtm)):
                error_message = f"The line with the expected values was not found for {nurse_name} in 'RawData' - PAYCODE: {expected_paycode}, STARTDTM: {expected_startdtm}"
                errors.append(error_message)

        # Check for cumulative errors and display them as a single assertion at the end
        self.assertEqual(errors, [], f"The following errors were found in the testcase:\n\n{', '.join(errors)}")


if __name__ == '__main__':
    unittest.main()
