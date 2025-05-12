import win32print
import json
import os
import subprocess

def listar_impresoras(show = True):
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    
    if show:
        print("Impresoras disponibles:")
        for i, printer in enumerate(printers, start=1):
            print(f"{i}. {printer[2]}")  # Nombre de la impresora

    return [printer[2] for printer in printers]

def filter_files(path):
    if not os.path.exists(path):
        raise FileNotFoundError("La ruta es inválida.")
    
    files = [file for file in os.listdir(path) if file.endswith('.pdf')]
    return files

def register_printer():
    impresoras_disponibles = listar_impresoras()

    while True:
        print("==============================================================")
        ans = input("\nSeleccione el numero de la impresora a usar: ").strip()
        try:
            select_print = impresoras_disponibles[int(ans)-1]
            break
        
        except IndexError as e:
            print(f"====== ERROR ====== {e}")
            print("El número no corresponde a ninguna impresora, vuelva a intentar...")

        except ValueError as e:
            print(f"====== ERROR ====== {e}")
            print("Debe ingresar un número, vuelva a intentar...")

    #Guardar impresora a usar.
    with open('config/printer.json', mode='w') as conf:
        json.dump({'name':select_print}, conf, indent = 4)


def printer_ghostscript(pdf_path, printer_name, number=1):
    """Envía un archivo PDF a la impresora de manera silenciosa con Ghostscript."""
    gs_command = [
        "C:\\Program Files\\gs\\gs10.05.1\\bin\\gswin64c.exe",
        "-dBATCH",
        "-dNOPAUSE",
        "-dQUIET",
        "-sDEVICE=mswinpr2",
        f"-sOutputFile=\\\\spool\\{printer_name}",
        f"-dNumCopies={number}",
        pdf_path
    ]
    
    try:
        subprocess.run(gs_command, check=True)
        return f"Impresión exitosa: {pdf_path}"
    except subprocess.CalledProcessError as e:
        return f"Error al imprimir {pdf_path}: {e}"
    
def ejecutar_impresor(dir_path, printer_name):
    """Ejecuta todo el flujo: filtra archivos y los envía a imprimir."""
    archivos_pdf = filter_files(dir_path)
    resultados = []

    print(f"Iniciando impresión en '{printer_name}'...")
    for archivo in archivos_pdf:
        pdf_path = os.path.join(dir_path, archivo)
        resultado = printer_ghostscript(pdf_path, printer_name)
        resultados.append({"archivo": archivo, "estado": resultado})
        print(resultado)

    # Guardar registro en JSON
    with open("config/log_impresion.json", "w") as log_file:
        json.dump(resultados, log_file, indent=4)

    print("Proceso de impresión completado. Log guardado en 'log_impresion.json'.")

carpeta_pdf = "files"

with open('config/printer.json', mode='r') as config:
    printer = json.load(config)
    printer_name = printer['printer']

ejecutar_impresor(carpeta_pdf, printer_name)
