import os
import win32print
import win32api
import logging

# Configurar el logging
logging.basicConfig(filename="errores.log", level=logging.ERROR)

def listar_impresoras_disponibles():
    # Obtiene una lista de impresoras disponibles
    impresoras = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    
    print("Impresoras disponibles:")
    for i, impresora in enumerate(impresoras):
        print(f"{i + 1}. {impresora[2]}")  # El índice 2 contiene el nombre de la impresora

def establecer_impresora(nombre_impresora):
    try:
        win32print.SetDefaultPrinter(nombre_impresora)
    except Exception as e:
        print(f"No se pudo establecer la impresora '{nombre_impresora}': {e}")
        logging.error(f"Error al establecer impresora: {e}")

def imprimir_archivos_en_carpeta(ruta_carpeta, impresora):
    try:
        establecer_impresora(impresora)
        
        # Lista los archivos en la carpeta
        archivos = os.listdir(ruta_carpeta)
        
        for archivo in archivos:
            if archivo.lower().endswith(".pdf"):  # Filtra solo PDFs
                ruta_archivo = os.path.join(ruta_carpeta, archivo)
                if os.path.isfile(ruta_archivo):  # Verifica que sea un archivo válido
                    print(f"Imprimiendo: {ruta_archivo}")
                    try:
                        win32api.ShellExecute(
                            0,
                            "print",
                            ruta_archivo,
                            None,
                            ".",
                            0
                        )
                    except Exception as e:
                        print(f"Error al imprimir {ruta_archivo}: {e}")
                        logging.error(f"Error con {ruta_archivo}: {e}")
    except Exception as e:
        print(f"Error general: {e}")
        logging.error(f"Error general: {e}")

# Especifica la impresora y la carpeta
listar_impresoras_disponibles()
nombre_impresora = input("Introduce el nombre de la impresora que deseas usar: ")
ruta_carpeta = "C:\\Users\\COMSITEC\\Desktop\\Impresor"
imprimir_archivos_en_carpeta(ruta_carpeta, nombre_impresora)

