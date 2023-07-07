import unittest
import pandas as pd

class ExcelTest(unittest.TestCase):
    def testCase1(self):

    #Descrition: Check last shift is present
    #File: Hannibal

        # Upload Excel file
        excel_file = 'Hannibal 4.15.23 SCHED.xlsx'
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

    def testCase3(self):

        #Descrition: Check Exe has the same data as last commit
        #File: TMMC WITHOUT SCHED

        # Loads the Excel file in a DataFrame
        df = pd.read_excel("TMMC W.E. 4.22.xlsx", sheet_name="OutputData")
        
        # Verify the number of rows
        self.assertEqual(len(df), 655-1, "Number of rows is not equal to 654")
        
        # Verify the existence of the column "NAME".
        self.assertIn("NAME", df.columns, "The column 'NAME' was not found'")
        
        # Gets the values of the column "NAME".
        names = df["NAME"]
        
        # Verify the values in the first and last row
        self.assertEqual(names.iloc[0], "Jo, Ahra", "The value in the first row of 'NAME' is not 'Jo, Ahra'.")
        self.assertEqual(names.iloc[-1], "Chekabab, Zahra", "The value in the last row of 'NAME' is not 'Chekabab, Zahra'.")


    def testCase3SCHED(self):

        #Descrition: Check Exe has the same data as last commit
        #File: TMMC WITH SCHED

        # Loads the Excel file in a DataFrame
        df = pd.read_excel("TMMC W.E. 4.22 SCHED.xlsx", sheet_name="OutputData")
        
        # Verify the number of rows
        self.assertEqual(len(df), 1310-1, "Number of rows is not equal to 1310")
        
        # Verify the existence of the column "NAME".
        self.assertIn("NAME", df.columns, "The column 'NAME' was not found")
        
        # Gets the values of the column "NAME".
        names = df["NAME"]
        
        # Verify the values in the first and last row
        self.assertEqual(names.iloc[0], "Yu, Ace", "The value in the first row of 'NAME' is not 'Yu, Ace'.")
        self.assertEqual(names.iloc[-1], "Chekabab, Zahra", "The value in the last row of 'NAME' is not 'Chekabab, Zahra'.")

    
        
if __name__ == '__main__':
    unittest.main()


