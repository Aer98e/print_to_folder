import win32print
import json
import os

def listar_impresoras(show = True):
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    
    if show:
        print("Impresoras disponibles:")
        for i, printer in enumerate(printers, start=1):
            print(f"{i}. {printer[2]}")  # Nombre de la impresora

    return [printer[2] for printer in printers]

def filter_files(path):
    if not os.path.exists(path):
        raise FileNotFoundError("La ruta es invalida.")
    
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
        json.dump({'printer':select_print}, conf, indent = 4)


filter_files('javier')
