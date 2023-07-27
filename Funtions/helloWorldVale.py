def imprimir_hello_world():
    print("Hello, World!\n")

def imprimir_test():
    print("test\n")

def suma_de_dos_Numeros(Numero1, Numero2):
    try:
        # Sumar los dos números ingresados
        suma = Numero1 + Numero2

        # Mostrar el resultado
        print(f"La suma de {Numero1} y {Numero2} es: {suma}")

    except ValueError:
        print("Error: Asegúrate de ingresar números válidos.")

# Ejemplo de cómo llamar a la función y pasar los valores como argumentos
numero1 = float(input("Ingresa el primer número: "))
numero2 = float(input("Ingresa el segundo número: "))
suma_de_dos_Numeros(numero1, numero2)


# Llamamos a la función para imprimir el mensaje
if __name__ == "__main__":
    imprimir_hello_world()
    imprimir_test()
    suma_de_dos_Numeros()
