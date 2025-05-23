import win32print
import json
import os
import subprocess

class ConfigManager():
    path = 'config'

    @classmethod
    def find_configuration(cls, name, safe_mode=False):
        files = get_files(cls.path, complete=True, endswith='json')

        select = next((file for file in files if name == os.path.splitext(os.path.basename(file))[0]), None)

        if select:
            return select
        
        if safe_mode:
            return None
        
        raise FileNotFoundError('El archivo no existe.')
    
    @classmethod
    def load_config(cls, name):
        select = cls.find_configuration(name, safe_mode=True)

        if select:
            with open(select, mode='r') as config:
                load=json.load(config)
            return load
        
        return None

    @classmethod
    def save_config(cls, name, json_dict, overwrite=True):
        # select = find_file(name, cls.path, safe_mode=True)
        path = os.path.join(cls.path, f'{name}.json')
        with open(path, mode='w') as config:
            json.dump(json_dict, config, indent=4)
    
    @classmethod
    def edit_config(cls, name, new_value):
        if not isinstance(new_value, dict):
            raise TypeError("El nuevo valor debe ser un diccionario.")
        
        select = cls.find_configuration(name, safe_mode=True)
        if not select:
            return None
        previous = cls.load_config(name)
        previous.update(new_value)
        cls.save_config(name, previous)
    
def test(fun):
    def wrapper(*args, **keyargs):
        print("_________________________________________________")
        print("Ejecutando función",fun.__name__)
        
        print("\tArgumentos: ")
        for arg in args:
            print(" -", arg)

        print("\tClave/Valor: ")
        for karg in keyargs:
            print(" -", karg)
        print("_________________________________________________")
    return wrapper

def get_printers(show = True):
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    
    if show:
        print("Impresoras disponibles:")
        for i, printer in enumerate(printers, start=1):
            print(f"{i}. {printer[2]}")  # Nombre de la impresora

    return [printer[2] for printer in printers]

def get_files(path, complete=True, endswith=None):
    if not os.path.exists(path):
        raise FileNotFoundError("La ruta es inválida.")
    
    if endswith and not '.' in endswith:
        endswith=f".{endswith}"

    files = [file for file in os.listdir(path) if os.path.isfile(os.path.join(path, file))]
    if endswith:
        files = list(filter(lambda file: file.endswith(endswith), files))
    if complete:
        files = list(map(lambda file:os.path.join(path, file), files))

    return files

def register_printer():
    printers_available = get_printers(show=True)
    attemps = 5

    for _ in range(attemps):
        print("==============================================================")
        ans = input("\nSeleccione el número de la impresora a usar: ").strip()
        try:
            select_print = printers_available[int(ans)-1]
            ConfigManager.edit_config('printer', {'name':select_print})
            return
        
        except (IndexError, ValueError) as e:
            print(f"\n====== ERROR ====== {e}")
            print("Entrada inválida, vuelva a intentar.")

    raise RuntimeError ("No se pudo establecer una impresora despues de varios intentos.")

@test
def printer_ghostscript(pdf_path, printer_name, number=1):
    """Envía un archivo PDF a la impresora de manera silenciosa con Ghostscript."""
    runner = ConfigManager.load_config('paths')['ghostscript']

    gs_command = [
        runner,
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

def _recognize(file: str):
    examine = file.split('-')
    if len(examine) > 1:
        return examine[-1].split('.')[-2].strip()
    return None

def ejecutar_impresor(dir_path):
    """Ejecuta todo el flujo: filtra archivos y los envía a imprimir."""
    configuration = ConfigManager.load_config('printer')
    printer = configuration['name']
    if not printer:
        register_printer()
        
    archivos_pdf = get_files(dir_path, endswith='pdf')
    log_print = []

    print(f"Iniciando impresión en '{printer}'...")
    for archivo in archivos_pdf:
        recognize = _recognize(archivo)
        profile = configuration['mode']
        number = configuration.get(profile, {}).get(recognize, 1)

        resultado = printer_ghostscript(archivo, printer, number)
        log_print.append({"archivo": archivo, "estado": resultado})
        print(resultado)

    # Guardar registro en JSON
    ConfigManager.save_config('log_impresion', log_print)

    print("Proceso de impresión completado. Log guardado en 'log_impresion.json'.")

def main():
    ejecutar_impresor('files')

if __name__ == "__main__":
    main()
