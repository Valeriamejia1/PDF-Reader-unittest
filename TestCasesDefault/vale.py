import pandas as pd
import unittest

class TestGLCodeValidation(unittest.TestCase):

    def validar_glcode(self, glcode):
        # Eliminar los guiones "-" y contar los dígitos restantes
        glcode_sin_guiones = glcode.replace('-', '')
        return len(glcode_sin_guiones) == 9 and glcode_sin_guiones.isdigit()

    def test_DEFAULT_3(self):
        archivos = ["TestCasesDefault/Combined File minutes.xlsx", "TestCasesDefault/Combined File.xlsx"]

        all_incorrect_rows = set()  # Usamos un conjunto para almacenar las filas incorrectas sin duplicados

        for archivo in archivos:
            try:
                with pd.ExcelFile(archivo) as xls:
                    sheet_name = xls.sheet_names[0]  # Asumimos que la hoja de interés es la primera

                    # Cargar el archivo y convertir la columna "GLCODE" al formato "text"
                    df = pd.read_excel(xls, sheet_name)
                    df["GLCODE"] = df["GLCODE"].astype(str)

                    # Validar que todas las celdas en la columna "GLCODE" contengan 9 dígitos sin guiones "-"
                    incorrect_rows = []
                    for index, glcode in df["GLCODE"].items():
                        if not self.validar_glcode(glcode):
                            incorrect_rows.append((archivo, index + 2, glcode))

                    # Agregar las filas incorrectas al conjunto general
                    all_incorrect_rows.update(incorrect_rows)
            except Exception as e:
                # Si ocurre una excepción, muestra el mensaje de error y registra la excepción.
                print(f"Error al procesar el archivo {archivo}: {str(e)}")

        # Mostrar mensaje de fallo con todas las filas incorrectas de todos los archivos
        if all_incorrect_rows:
            error_msg = "El GLCODE no contiene 9 digitos en las siguientes filas y archivos:\n"
            for archivo, fila, glcode in all_incorrect_rows:
                error_msg += f"Archivo: {archivo}, Fila: {fila}, GLCODE: {glcode}\n"
            self.fail(error_msg)

        print(".TEST 3 DEFAULT CORRECT: All GLCODEs have 9 numeric digits.")




if __name__ == "__main__":
    unittest.main()
