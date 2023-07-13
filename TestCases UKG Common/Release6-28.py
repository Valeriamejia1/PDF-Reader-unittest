import pandas as pd
import unittest

class TestCase6(unittest.TestCase):
    def test_comments_column(self):
        # Cargar el archivo de Excel
        df = pd.read_excel("Martin ppe 4.22.23.xlsx")

        # Verificar si la columna "Comments" contiene datos
        if "Comments" not in df.columns:
            self.fail("No 'Comments' column found in file.")

        comments_column = df["Comments"]
        if comments_column.dropna().empty:
            self.fail("There is not data in the column 'Comments'.")
        else:
            print("The 'Comments' column contains datas.")

if __name__ == '__main__':
    unittest.main()
