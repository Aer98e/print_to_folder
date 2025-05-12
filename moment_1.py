import win32print
import json
import os
import subprocess

def get_printers(show = True):
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    
    if show:
        print("Impresoras disponibles:")
        for i, printer in enumerate(printers, start=1):
            print(f"{i}. {printer[2]}")  # Nombre de la impresora

    return [printer[2] for printer in printers]

def get_files(path, complete=True, endswith='pdf'):
    if not os.path.exists(path):
        raise FileNotFoundError("La ruta es inválida.")
    
    origin_dir = os.getcwd()
    os.chdir(path)
    files = [file for file in os.listdir() if os.path.isfile(file)]
    if endswith:
        files = list(filter(lambda file: file.endswith(f'.{endswith}'), files))
    if complete:
        files = list(map(lambda file:os.path.join(path, file), files))
    
    os.chdir(origin_dir)

    return files

def register_printer(path:str, diction:dict=None):
    printers_available = get_printers()

    while True:
        print("==============================================================")
        ans = input("\nSeleccione el número de la impresora a usar: ").strip()
        try:
            select_print = printers_available[int(ans)-1]
            break
        
        except IndexError as e:
            print(f"\n====== ERROR ====== {e}")
            print("El número no corresponde a ninguna impresora, vuelva a intentar...")

        except ValueError as e:
            print(f"\n====== ERROR ====== {e}")
            print("Debe ingresar un número, vuelva a intentar...")

    if not diction:
        diction = {'name': select_print}
    else:
        diction['name'] = select_print

    with open(path, mode='w') as conf:
        json.dump(diction, conf, indent = 4)


def printer_ghostscript(pdf_path, printer_name, number=1):
    print(f"imprimiendo {pdf_path} en {printer_name} -{number} veces.")
    return None
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
    archivos_pdf = get_files(dir_path)
    log_print = []

    print(f"Iniciando impresión en '{printer_name}'...")
    for archivo in archivos_pdf:
        pdf_path = os.path.join(dir_path, archivo)
        resultado = printer_ghostscript(pdf_path, printer_name)
        log_print.append({"archivo": archivo, "estado": resultado})
        print(resultado)

    # Guardar registro en JSON
    with open("config/log_impresion.json", "w") as log_file:
        json.dump(log_print, log_file, indent=4)

    print("Proceso de impresión completado. Log guardado en 'log_impresion.json'.")

def main():
    carpeta_pdf = "files"

    with open('config\\printer.json', mode='r') as config:
        reg_config = json.load(config)

    files = get_files(carpeta_pdf, complete=True, endswith='txt')
    for file in files:
        examine = file.split('-')

        if len(examine) > 1:
            key_word = examine[-1].split('.')[-2].strip()
            if key_word in reg_config[reg_config['mode']]:
                petition = reg_config[reg_config['mode']][key_word]
            else:
                petition = 1
        
        printer_ghostscript(file, reg_config['name'], petition)







    



if __name__ == "__main__":
    main()
