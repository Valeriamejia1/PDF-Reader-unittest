import unittest
import pandas as pd
import numpy as np
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)


class TestExcel(unittest.TestCase):

    def test_API_1(self):
       
        # Description TestCase: Output Data Should remove shifts with Paycode LCUP, SHDIF and LUNCH
        # File: TMMC

        # Read Excel file and select the sheet "OutputData"
        data_frame = pd.read_excel("OUTPUT API/TMMC W.E. 4.22.xlsx", sheet_name="OutputData")

        # Obtain the column
        Column = data_frame["PAYCODE"]

        # Checked the expected Values
        expectedValues = ["LUNCH", "LCUP", "SHDIF"]

        missingValues = []

        # if an expected value is not found in column.values, it is added to the missing_values list.
        for value in expectedValues:
            if value in Column.values:
                missingValues.append(value)

        self.assertListEqual(missingValues, [], f"The following values are not present in the PAYCODE column in Excel TMMC file: {missingValues}")

        if not missingValues:
            
            print(".TEST 1 API CORRECT: The file TMMC W.E. 4.22 does not contain data 'LUNCH', 'LCUP', 'SHDIF' in the column 'PAYCODE'")

    def test_API_2(self):
    
        # Description TestCase: Remove SCHED shifts when necessary
        # File: DELTA

        data_frame = pd.read_excel("OUTPUT API/Delta Health 4.15.23.xlsx", sheet_name="OutputData")

        column = data_frame["PAYCODE"]

        expectedValues = ["SCHED"]

        for value in column:
            self.assertNotIn(value, expectedValues, f"The value '{value}' is present in the PAYCODE column in the Delta Health 4.15.23 file.")

        print("TEST 2 API CORRECT: The value SCHED is NOT present in the PAYCODE column in the Delta Health 4.15.23 file.")
    
    def test_API_3(self):

        #Description TestCase: Remove SCHED shifts when is neccesary
        #File: TMMC
        data_frame = pd.read_excel("OUTPUT API/TMMC W.E. 4.22.xlsx", sheet_name="OutputData")

        Column = data_frame["PAYCODE"]

        expectedValues = ["SCHED"]

        for value in Column:
            self.assertNotIn(value, expectedValues, f"The value '{value}' is present in the PAYCODE column in the TMMC file.")

        print("TEST 3 API CORRECT: The value SCHED is NOT present in the PAYCODE column in the TMMC W.E. 4.22 file.")
    
    def test_API_4(self):
        # Carga el archivo de Excel en un DataFrame
        df = pd.read_excel('OUTPUT API/TMMC W.E. 4.22.xlsx', sheet_name='OutputData')
        
        # Filtra las filas donde el shift de PAYCODE está vacío pero el de HOURS no
        empty_hours = df[(df['PAYCODE'].notnull()) & (df['HOURS'].isnull())]
        
        # Filtra las filas donde el shift de PAYCODE no está vacío pero el de HOURS sí
        empty_paycode = df[(df['PAYCODE'].isnull()) & (df['HOURS'].notnull())]
        
        # Comprueba si hay filas que cumplan con los filtros
        if not empty_hours.empty:
            failed_rows = empty_hours.index + 2
            self.fail("PAYCODE shift is not empty but HOURS shift is empty in rows {}".format(failed_rows))

        if not empty_paycode.empty:
            failed_rows = empty_paycode.index + 2
            self.fail("The PAYCODE shift is empty but the HOURS shift is not in the rows. {}".format(failed_rows))

        print(".TEST 4 API CORRECT: If no data in some shift PAYCODES in the HOURS shift there is no data either., in the file TMMC W.E. 4.22.xlsx")

    def test_API_5(self):

        # Loads the Excel file into a DataFram
        df = pd.read_excel('OUTPUT API/API Empty.xlsx', header=None)

        # Gets the number of rows with data beyond the headers
        num_data_rows = len(df) - 1 

        # Check if there are additional rows with data
        if num_data_rows > 0:

            self.fail("There are {} additional rows with data in the Excel file".format(num_data_rows))
        
        print(".TEST 5 API CORRECT: The API Empty.xlsx file not contains additional rows to the header")
    
    def test_API__6(self):

    #Descrition: Check last shift is present
    #File: Hannibal

        # Upload Excel file
        excel_file = 'OUTPUT API/Hannibal 4.15.23 SCHED.xlsx'
        df = pd.read_excel(excel_file, sheet_name='OutputData')

        # Specify the search criteria
        name = 'WYCOFF, JENNA'
        date = '04/07/2023'
        paycode = 'REG'
        hours = 13.10

        # Filter the DataFrame based on the following criteria
        filtered_df = df[
            (df['NAME'] == name) &
            (df['DATE'] == date) &
            (df['PAYCODE'] == paycode) &
            (df['HOURS'] == hours)
        ]

        # Check if matching rows were found
        if filtered_df.empty:
            error_message = f"No matching row was found in the Excel file for the following values:\n\n" \
                            f"Name: {name}\n" \
                            f"Date: {date}\n" \
                            f"Paycode: {paycode}\n" \
                            f"Hours: {hours}\n"
            self.fail(error_message)
        else:
            self.assertEqual(len(filtered_df), 1, 'Multiple matches found in Excel file.')

        print(".TEST 6 API CORRECT: Checked that the last line of the file TestCasesAPI/Hannibal 4.15.23 SCHED.xlsx is still for WYCOFF, JENNA with the same data")
    
    def test_API_7(self):

        #Descrition: Validate Output has all nurses
        #File: DELTA modified
        data_frame = pd.read_excel("OUTPUT API/Delta Health 4.15.23 SCHED.xlsx", sheet_name="OutputData")
        Column = data_frame["NAME"]
        expectedValues = ["Hunter, Angelique","Halums, Brittney","Cross, Destin","Radford, Gladys","Hale, Shannon","Kelly, Joby", "Lowe, Sherrie", "Lewis, Susan", "Towery, Brittany"]
        #Added a missing_values list to store the values that were not found in the "NAME" column.
        missingValues = []

        #if an expected value is not found in column.values, it is added to the missing_values list.
        for value in expectedValues:
            if value not in Column.values:
                missingValues.append(value)
        self.assertFalse(missingValues, f"The following values are not present in the NAME column in Excel Delta file: {missingValues}")

        print("TEST 7 API CORRECT: Se encontro al menos una fila para: Hunter, Angelique,Halums, Brittney,Cross, Destin,Radford, Gladys,Hale, Shannon,Kelly, Joby, Lowe, Sherrie, Lewis, Susan, Towery, Brittany in file Delta Health 4.15.23 SCHED")

    def test_API_9(self):
        # Description: Validate Output Out time from TMMC as it takes the out time
        # File: TMMC WITH SCHED

        # Upload Excel file
        df = pd.read_excel("OUTPUT API/TMMC W.E. 4.22 SCHED.xlsx", sheet_name="OutputData")

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

        print("TEST 9 API CORRECT:Data for Chekabab, Zahra and for Maldonado, Marleny were found in the file TMMC W.E. 4.22 SCHED in OutputData sheet")

    def test_API_10(self):
        # Descrition: Check RAW data Out has the same hour as output
        # File: TMMC WITH SCHED

        # Upload Excel file
        df_output = pd.read_excel("OUTPUT API/TMMC W.E. 4.22 SCHED.xlsx", sheet_name="OutputData")
        df_raw = pd.read_excel("OUTPUT API/TMMC W.E. 4.22 SCHED.xlsx", sheet_name="RawData")

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
            print("TEST 10 API CORRECT: Data for Chekabab, Zahra and for Maldonado, Marleny were found in the file TMMC W.E. 4.22 SCHED in OutputData and RawData sheet")

    def test_API_11(self):

        #Descrition: Check Exe has the same data as last commit
        #File: TMMC WITH SCHED

        # Loads the Excel file in a DataFrame
        df = pd.read_excel("OUTPUT API/TMMC W.E. 4.22 SCHED.xlsx", sheet_name="OutputData")
        
        # Verify the number of rows
        self.assertEqual(len(df), 1349-1, "Number of rows is not equal to 1349")
        
        # Verify the existence of the column "NAME".
        self.assertIn("NAME", df.columns, "The column 'NAME' was not found")
        
        # Gets the values of the column "NAME".
        names = df["NAME"]
        
        # Verify the values in the first and last row
        self.assertEqual(names.iloc[0], "Yu, Ace", "The value in the first row of 'NAME' is not 'Yu, Ace'.")
        self.assertEqual(names.iloc[-1], "Chekabab, Zahra", "The value in the last row of 'NAME' is not 'Chekabab, Zahra'.")

        print("TEST 11 API CORRECT: Checked that the file TestCasesAPI/TestCasesAPI/TMMC W.E. 4.22 SCHED.xlsx still has the same number of rows and that the first and last row are the same.")

    #Methods required for API_TEST_12 and API_TEST_13

    def compare_excel_files(self, original_file, new_file):
        # Loads the original Excel file in a DataFrame
        original_df = pd.read_excel(original_file, sheet_name="OutputData")

        # Loads the new Excel file in a DataFrame
        new_df = pd.read_excel(new_file, sheet_name="OutputData")

        # Verify that the DataFrames are equal
        self.assertTrue(
            original_df.equals(new_df),
            self.generate_difference_message(original_df, new_df, original_file, new_file),
        )

    def generate_difference_message(self, original_df, new_df, original_file, new_file):
        original_df = original_df.fillna("NA")
        new_df = new_df.fillna("NA")

        diff_df = original_df != new_df
        diff_indices = diff_df.any(axis=1)

        diff_message = f"Differences found between {original_file} and {new_file}:\n"
        for index, row in diff_df[diff_indices].iterrows():
            diff_message += f"Row {index+2}:\n"
            for column, value in row.items():
                if value:
                    original_value = original_df.at[index, column]
                    new_value = new_df.at[index, column]
                    diff_message += f"  - Column '{column}': Original='{original_value}', New='{new_value}'\n"

        return diff_message
    
    #Descrition: Compare original files with new file versions of the same PDF file
    #Files: Test Dawson, Kathleen //  Mattox, Kyle  // Hannibal 4.15.23

    def test_API_12_1(self):
        self.compare_excel_files("TestCasesAPI/ORIG Files/Dawson, Kathleen ORIG.xlsx", "OUTPUT API/Dawson, Kathleen.xlsx")
        self.assertTrue(True)
        print("TEST 12.1 API CORRECT: The Dawson, Kathleen.xlsx data match the original version.")

    def test_API_12_2(self):
        self.compare_excel_files("TestCasesAPI/ORIG Files/Mattox, Kyle ORIG.xlsx", "OUTPUT API/Mattox, Kyle.xlsx")
        self.assertTrue(True)
        print("TEST 12.2 API CORRECT:The Mattox, Kyle.xlsx data match the original version.")

    def test_API_13(self):
        self.compare_excel_files("TestCasesAPI/ORIG Files/Hannibal 4.15.23 ORIG.xlsx", "OUTPUT API/Hannibal 4.15.23.xlsx")
        self.assertTrue(True)
        print("TEST 12.3 API CORRECT:The Hannibal 4.15.23.xlsx data match the original version.")
        

if __name__ == '__main__':
    unittest.main()
