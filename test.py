import unittest
import pandas as pd
from datetime import datetime

class TestExcel(unittest.TestCase):
    def testCase9(self):

        # Validate Output Out time from TMMC as it takes the out time
        # Archivo: TMMC without SCHED

        data_frame = pd.read_excel("TMMC W.E. 4.22.xlsx", sheet_name="OutputData")
        nurse_column = data_frame["NAME"]
        startdtm_column = pd.to_datetime(data_frame["STARTDTM"])
        enddtm_column = pd.to_datetime(data_frame["ENDDTM"])
        paycode_column = data_frame["PAYCODE"]

        nurse_name_1 = "Chekabab, Zahra"
        expected_value_enddtm_1 = datetime(2023, 4, 18, 7, 39)
        expected_value_startdtm_1 = datetime(2023, 4, 17, 18, 58)

        nurse_name_2 = "Maldonado, Marleny"
        expected_value_enddtm_2 = datetime(2023, 4, 20, 7, 26)
        expected_value_startdtm_2 = datetime(2023, 4, 19, 18, 55)

        expected_paycode = "REG"
        errors = []

        found_nurse_1 = False
        found_nurse_2 = False
        for nurse, startdtm, enddtm, paycode in zip(nurse_column, startdtm_column, enddtm_column, paycode_column):
            if not found_nurse_1 and nurse == nurse_name_1 and paycode == expected_paycode:
                if enddtm != expected_value_enddtm_1:
                    errors.append(f"The value '{expected_value_enddtm_1}' is not present in ENDDTM for the nurse '{nurse_name_1}' in the TMM file.")
                if startdtm != expected_value_startdtm_1:
                    errors.append(f"The value '{expected_value_startdtm_1}' is not present in STARTDTM for the nurse '{nurse_name_1}' in the TMM file.")
                found_nurse_1 = True

            if not found_nurse_2 and nurse == nurse_name_2 and paycode == expected_paycode:
                if enddtm != expected_value_enddtm_2:
                    errors.append(f"The value '{expected_value_enddtm_2}' is not present in ENDDTM for the nurse '{nurse_name_2}' in the TMM file.")
                if startdtm != expected_value_startdtm_2:
                    errors.append(f"The value '{expected_value_startdtm_2}' is not present in STARTDTM for the nurse '{nurse_name_2}' in the TMM file.")
                found_nurse_2 = True

            if found_nurse_1 and found_nurse_2:
                break

        if errors:
            self.fail("\n".join(errors))
