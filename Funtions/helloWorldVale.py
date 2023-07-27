def imprimir_hello_world():
    print("Hello, World!\n")

def imprimir_test():
    print("test\n")

def suma_de_dos_Numeros(Numero1, Numero2):
    try:
        Numero1 = float(Numero1)
        Numero2 = float(Numero2)

        suma = Numero1 + Numero2
        print(f"La suma de {Numero1} y {Numero2} es: {suma}")

    except ValueError:
        print("Error: Asegúrate de ingresar números válidos.")

# Llamamos a la función para imprimir el mensaje
if __name__ == "__main__":
    imprimir_hello_world()
    imprimir_test()


