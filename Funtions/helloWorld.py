import os
import sys

def suma_de_dos_Numeros(Numero1, Numero2):
    try:
        # Convertir los valores ingresados a float
        Numero1 = float(Numero1)
        Numero2 = float(Numero2)

        # Sumar los dos números ingresados
        suma = Numero1 + Numero2

        # Mostrar el resultado
        print(f"La suma de {Numero1} y {Numero2} es: {suma}")

    except ValueError:
        print("Error: Asegúrate de ingresar números válidos.")

if __name__ == "__main__":
    # Verificar si se proporcionaron suficientes argumentos
    if len(sys.argv) == 3:
        # Obtener los valores de los números de los argumentos
        Numero1 = sys.argv[1]
        Numero2 = sys.argv[2]

        # Llamar a la función con los valores proporcionados
        suma_de_dos_Numeros(Numero1, Numero2)
    else:
        print("Error: Se esperan dos argumentos. Ejemplo: python mi_script.py 5 10")
