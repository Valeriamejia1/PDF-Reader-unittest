def imprimir_hello_world():
    print("Hello, World!\n")

def imprimir_test():
    print("test\n")

def suma_de_dos_numeros():
    try:
        # Pedir al usuario que ingrese los dos números
        Numero1 = float(input("Ingresa el primer número: "))
        Numero2 = float(input("Ingresa el segundo número: "))

        # Sumar los dos números ingresados
        suma = numero1 + numero2

        # Mostrar el resultado
        print(f"La suma de {numero1} y {numero2} es: {suma}")

    except ValueError:
        print("Error: Asegúrate de ingresar números válidos.")

# Llamar a la función para ejecutarla





# Llamamos a la función para imprimir el mensaje
if __name__ == "__main__":
    imprimir_hello_world()
    imprimir_test()
    suma_de_dos_numeros()

