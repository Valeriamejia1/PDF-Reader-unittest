import unittest
import pandas as pd
import numpy as np

class ExcelTestCase(unittest.TestCase):
    def testcase9(self):
        # Cargar el archivo Excel
        df = pd.read_excel("TMMC W.E. 4.22 SCHED.xlsx", sheet_name="OutputData")

        # Formatear las columnas de fecha y hora para que coincidan con el formato esperado
        df["STARTDTM"] = pd.to_datetime(df["STARTDTM"], format="%m/%d/%Y %H:%M")
        df["ENDDTM"] = pd.to_datetime(df["ENDDTM"], format="%m/%d/%Y %H:%M")
        
        # Validar datos para la enfermera "Chekabab, Zahra"
        nurse_name = "Chekabab, Zahra"
        
        nurse_data = df[df["NAME"] == nurse_name]
        expected_data = [
            ("REG", pd.to_datetime("2023-04-17 18:58:00"), pd.to_datetime("2023-04-18 07:39:00")),
            ("SCHED", pd.to_datetime("2023-04-18 19:01:00"), pd.to_datetime("2023-04-19 07:34:00"))
        ]
        
        for expected_paycode, expected_startdtm, expected_enddtm in expected_data:
            self.assertTrue(any((nurse_data["PAYCODE"] == expected_paycode) & (nurse_data["STARTDTM"] == expected_startdtm) & (nurse_data["ENDDTM"] == expected_enddtm)), f"No se encontró la línea con los valores esperados para {nurse_name} - PAYCODE: {expected_paycode}, STARTDTM: {expected_startdtm}, ENDDTM: {expected_enddtm}")
        
        # Validar datos para la enfermera "Maldonado, Marleny"
        nurse_name = "Maldonado, Marleny"
        
        nurse_data = df[df["NAME"] == nurse_name]
        expected_data = [
            ("REG", pd.to_datetime("2023-04-19 18:55:00"), pd.to_datetime("2023-04-20 07:26:00")),
            ("SCHED", pd.to_datetime("2023-04-24 19:00:00"), np.nan)
        ]
        
        for expected_paycode, expected_startdtm, expected_enddtm in expected_data:
            if pd.isnull(expected_enddtm):
                self.assertTrue(any((nurse_data["PAYCODE"] == expected_paycode) & (nurse_data["STARTDTM"] == expected_startdtm) & pd.isnull(nurse_data["ENDDTM"])), f"No se encontró la línea con los valores esperados para {nurse_name} - PAYCODE: {expected_paycode}, STARTDTM: {expected_startdtm}, ENDDTM: NaN")
            else:
                self.assertTrue(any((nurse_data["PAYCODE"] == expected_paycode) & (nurse_data["STARTDTM"] == expected_startdtm) & (nurse_data["ENDDTM"] == expected_enddtm)), f"No se encontró la línea con los valores esperados para {nurse_name} - PAYCODE: {expected_paycode}, STARTDTM: {expected_startdtm}, ENDDTM: {expected_enddtm}")


if __name__ == '__main__':
    unittest.main()
