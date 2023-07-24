import unittest
import pandas as pd

class TestGLCode(unittest.TestCase):
    import unittest
import pandas as pd

class TestGLCode(unittest.TestCase):
    def test_glcode_match(self):
        file1 = "TestCasesDefault/Time Detail_July152022 minutes.xlsx"
        file2 = "TestCasesDefault/Time Detail_July152022.xlsx"
        name_to_find = "Anderson, Kasey"
        expected_glcode = ["34006510", "4900-20-40"]

        # Leer los archivos Excel
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)

        # Filtrar por el valor de "NAME"
        filtered_df1 = df1[df1["NAME"] == name_to_find]
        filtered_df2 = df2[df2["NAME"] == name_to_find]

        # Almacenar los errores encontrados
        errors = []

        # Verificar que los valores de "GLCODE" coincidan con los esperados
        for index, row in filtered_df1.iterrows():
            if row["GLCODE"] not in expected_glcode:
                errors.append(f"File: {file1}, Row: {index + 2}")

        for index, row in filtered_df2.iterrows():
            if row["GLCODE"] not in expected_glcode:
                errors.append(f"File: {file2}, Row: {index + 2}")

        # Si hay errores, hacer que el unittest falle con assert.fail()
        if errors:
            assert False, "\n".join(["ERROR: " + error for error in errors])

        # Si no hay errores, el unittest pasa correctamente
        print("TEST 7 DEFAULT CORRECT: The GLCODE of Anderson, Kasey match the expected value")

if __name__ == "__main__":
    unittest.main()


if __name__ == "__main__":
    unittest.main()
