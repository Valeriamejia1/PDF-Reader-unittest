import unittest
import pandas as pd

class TestExcelData(unittest.TestCase):
    def testcase9(self):
        # Validate Output Out time from TMMC as it takes the out time
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


if __name__ == "__main__":
    unittest.main()
