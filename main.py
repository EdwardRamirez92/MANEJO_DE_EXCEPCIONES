import pandas as pd
import openpyxl as opp
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# ---------------- FUNCIONES ---------------- #
def seleccionar_archivo_excel():
    tk.Tk().withdraw()
    archivo = filedialog.askopenfilename(
        title="Selecciona archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
    )
    if archivo:
        df = pd.read_excel(archivo)
        print("Archivo cargado correctamente:")
        print(df.head())
    else:
        print("No se seleccionó ningún archivo")

def mostrar_hojas_excel():
    wb = opp.load_workbook("BASE_MATERIALES_APPSHEETS.xlsx")
    hojas = wb.sheetnames
    messagebox.showinfo("Hojas del Excel", f"El libro contiene las hojas:\n{hojas}")

def leer_columna_A():
    df = pd.read_excel("BASE_MATERIALES_APPSHEETS.xlsx", sheet_name="Inspecciones")
    columna_a = df["A"].tolist()
    print("Columna A:", columna_a)

def crear_archivo(nombre):
    try:
        with open(nombre, 'x') as f:
            print(f'Se creó el archivo: {nombre}')
    except FileExistsError:
        print(f"El archivo {nombre} ya existe")

def escribir_archivo(nombre, contenido):
    with open(nombre, 'w') as f:
        f.write(contenido)
        print(f"Se escribió contenido en: {nombre}")

def leer_archivo(nombre):
    with open(nombre, 'r') as f:
        print(f.readlines())

def anexar_archivo(nombre, lineas):
    with open(nombre, 'a') as f:
        f.writelines(lineas)
        print("Información anexada")

def eliminar_archivo(nombre):
    if os.path.exists(nombre):
        os.remove(nombre)
        print(f"Archivo eliminado: {nombre}")
    else:
        print("El archivo no existe")

def manejo_excepciones():
    try:
        x = int(input("Dame el valor de x: "))
        if x == 0:
            raise Exception("La variable x es 0")
    except Exception as e:
        print(f"Ocurrió un error: {e}")
    else:
        print("x es distinto de 0")
    finally:
        print("Validación finalizada")

def abrir_archivo_con_manejo(nombre):
    archivo = None
    try:
        archivo = open(nombre)
        print(f"Contenido:\n{archivo.read()}")
    except Exception as e:
        print(f"Error al abrir archivo: {e}")
    finally:
        if archivo:
            archivo.close()
            print("Archivo cerrado correctamente")
        else:
            print("Archivo no inicializado")

# ---------------- MENÚ PRINCIPAL ---------------- #
def main():
    while True:
        print("""
        --- MENU DE OPCIONES ---
        1. Seleccionar archivo Excel
        2. Mostrar hojas del Excel
        3. Leer columna A de hoja Inspecciones
        4. Crear archivo
        5. Escribir en archivo
        6. Leer archivo
        7. Anexar a archivo
        8. Eliminar archivo
        9. Validar Excepciones
       10. Abrir archivo con manejo de errores
       11. Salir
        """)
        opcion = input("Elige una opción: ")

        if opcion == '1': seleccionar_archivo_excel()
        elif opcion == '2': mostrar_hojas_excel()
        elif opcion == '3': leer_columna_A()
        elif opcion == '4': crear_archivo(input("Nombre: "))
        elif opcion == '5': escribir_archivo(input("Nombre: "), input("Contenido: "))
        elif opcion == '6': leer_archivo(input("Nombre: "))
        elif opcion == '7': anexar_archivo(input("Nombre: "), ['\nLínea uno\n', 'Línea dos\n'])
        elif opcion == '8': eliminar_archivo(input("Nombre: "))
        elif opcion == '9': manejo_excepciones()
        elif opcion == '10': abrir_archivo_con_manejo(input("Nombre: "))
        elif opcion == '11': break
        else: print("Opción no válida")

if __name__ == '__main__':
    main()
