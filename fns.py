from datetime import datetime, timedelta
import requests
import json
import os
from decouple import config
import pandas as pd
from lxml import etree
import tkinter as tk
from tkinter.filedialog import askopenfilename
from PIL import Image, ImageTk
from tkinter import filedialog, messagebox
import sys
from tkinter import filedialog, messagebox
from datetime import datetime

def validar_columnas(df):
    """
    PENDIENTE: 
    Pretende validar que la informacion de el excel sea la esperada.
    """
    # Lista de columnas requeridas
    columnas_requeridas = [
        "NumeroPermisoCREProveedor", "Diaareportar", "ProductoId", "SubProductoId", 
        "TipoMov", "CostoFlete", "VolumenVendido", "PrecioVenta", "VolumenComprado", 
        "PrecioCompra", "PrecioDescuentoIncluido", "PermisoTransportista", 
        "NumeroPermisoCRECliente", "TipoDescuentoId", "Entidad", "Municipio", 
        "RazonSocial", "RFC", "SectorEconomico", "TipoCliente"
    ]
    
    # Lista de las columnas faltantes
    columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]
    
    if columnas_faltantes:
        print("Las siguientes columnas faltan o están mal escritas en el DataFrame:")
        for col in columnas_faltantes:
            print(f"- {col}")
        return False
    else:
        print("Todas las columnas requeridas están presentes.")
        return True


def cargar_excel():
    # Crear la ventana principal para configurar el ícono
    ventana = tk.Tk()
    ventana.withdraw()  # Oculta la ventana principal

    # Configurar el ícono (usando archivo .ico o .png según tu sistema)
    try:
        ventana.iconbitmap(config('ICONO_VENTANA'))  # Reemplaza con la ruta de tu archivo .ico
    except:
        # Si el sistema no soporta iconbitmap (como en algunas distribuciones Linux), usa iconphoto con .png
        imagen = Image.open(config('PNG_VENTANA'))  # Reemplaza con la ruta de tu archivo .png
        icono = ImageTk.PhotoImage(imagen)
        ventana.iconphoto(True, icono)
    
    # Configurar la ventana de diálogo
    archivo_path = askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")],  # Solo mostrar archivos de Excel
        title=config('SELECT_FILE_WINDOW')
    )
    ventana.destroy()  # Cierra la ventana de Tkinter una vez seleccionado el archivo

    if archivo_path:  # Si se selecciona un archivo
        try:
            df = pd.read_excel(archivo_path)
            print("Archivo cargado correctamente.")
            return df
        except PermissionError:
            msg = f"Error: El archivo seleccionado está abierto. Por favor, ciérralo e inténtalo de nuevo."
            print(msg)
            msg_txt(msg)
            msg_exit(msg)
        except Exception as e:
            msg = f"Error al cargar el archivo: {e}"
            print(msg)
            msg_txt(msg)
            msg_exit(msg)
    else:
        msg = f"No se seleccionó ningún archivo."
        print(msg)
        msg_txt(msg)
        msg_exit(msg)

def pandas_consola():
    """ Nos permite agregar opciones a pandas para visualizar en la consola.
        Se puede omitir su uso cuando ya no se encuentre en un entorno de pruebas."""
    # Configurar Pandas para mostrar más datos
    pd.set_option('display.max_rows', None)  # Mostrar todas las filas al imprimir en consola. 
    pd.set_option('display.max_columns', None)  # Mostrar todas las columnas al imprimir en consola. 
    pd.set_option('display.width', None)  # Limita el ancho de la visualización para no desplazar la pantalla. 

def fechas_inicio_fin(fecha=None, valor=0):
    # Si no se proporciona una fecha, se utiliza la fecha actual
    hoy = datetime.strptime(fecha, "%Y-%m-%d") if fecha else datetime.now()
    # Encontrar el jueves anterior más cercano (sin incluir hoy)
    dias_retroceso_jueves = ((hoy.weekday() - 3) % 7) + 7  # 3 representa el jueves
    jueves_anterior = hoy - timedelta(days=dias_retroceso_jueves)
    # Calcular el viernes anterior a ese jueves
    viernes_anterior = jueves_anterior - timedelta(days=6)
    
    if valor == 1:
        # Generar lista de todas las fechas entre viernes y jueves
        dias = [(viernes_anterior + timedelta(days=i)).strftime('%Y-%m-%d') 
                for i in range((jueves_anterior - viernes_anterior).days + 1)]  
        return dias
    else:
        # Devolver solo viernes y jueves
        fechainicial = viernes_anterior.strftime('%Y-%m-%d')
        fechafinal = jueves_anterior.strftime('%Y-%m-%d')
        return fechainicial, fechafinal


def modificacion(ruta_archivo):
    try:
        timestamp = os.path.getmtime(ruta_archivo)
        fecha_modificacion = datetime.fromtimestamp(timestamp)
        fecha_modificacion = fecha_modificacion.strftime('%d-%m-%Y')
        return fecha_modificacion
    except FileNotFoundError:
        msg = f"El archivo no existe en la ruta especificada."
        print(msg)
        msg_txt(msg)
        msg_pass(msg)

def api_catalogoproducto():
    r = f"{config('PATH_PRODUCTOS_JSON')}.json"
    url = config('URL_PRODUCTOS')
    try:
        res = requests.get(url, verify=False)
        if res.status_code == 200:
            productos = res.json()
            os.makedirs(config('DIR_CATALOGOS'), exist_ok=True)
            with open(r, "w", encoding="utf-8") as json_file:
                json.dump(productos, json_file, ensure_ascii=False, indent=4)
        else:
            m = modificacion(r)
            msg = f"No se pudo obtener el catálogo de productos.\n Archivo actualizado por ultima vez el: {m}"
            print(msg)
            msg_txt(msg)
            msg_pass(msg)
    except Exception as e:
        print(f"Error: {e}")


def api_catalogosubproductos():
    e = config('PRODUCT_ID')
    e_l = [int(i) for i in e.split(",") if i]
    msg = []
    for _e in e_l:
        r = f"{config('PATH_SUBPRODUCTOS_JSON')}{_e}.json"
        url = f"{config('URL_SUBPRODUCTOS')}{_e}"
        try:
            res = requests.get(url, verify=False)
            if res.status_code == 200:
                productos = res.json()
                os.makedirs(config('DIR_CATALOGOS'), exist_ok=True)
                with open(r, "w", encoding="utf-8") as json_file:
                    json.dump(productos, json_file, ensure_ascii=False, indent=4)
            else:
                m = modificacion(r)
                msg.append(f"No se pudo obtener el catálogo de productos para el producto: {_e}.\n Archivo actualizado por ultima vez el: {m} ")
        except Exception as e:
            print(f"Error: {e}")
    

def validar_xml_con_xsd(xml_str, xsd_path):
    # Cargar y compilar el esquema XSD
    with open(xsd_path, 'rb') as xsd_file:
        esquema_xsd = etree.XMLSchema(etree.parse(xsd_file))
    
    # Parsear la cadena XML
    try:
        xml_tree = etree.fromstring(xml_str.encode('utf-8'))
    except etree.XMLSyntaxError as e:
        msg = f"Error de sintaxis en el XML: {str(e)}"
        print(msg)
        msg_txt(msg)
        msg_pass(msg)
        return False
    
    # Validar el XML contra el esquema XSD
    if esquema_xsd.validate(xml_tree):
        msg = f"El XML es válido según el esquema XSD."
        print(msg)
        msg_txt(msg)
        msg_pass(msg)
        return True
    else:
        msg = f"El XML NO es válido según el esquema XSD."
        msg_2 = f"Seleccione una carpeta para guardar un resumen del error."
        print(msg)
        msg_txt(msg)
        msg_pass(msg)
        print(msg_2)
        msg_pass(msg_2)
        # Opcional: imprimir errores
        err = []
        for error in esquema_xsd.error_log:
            msg = f"Error: {error.message}, Línea: {error.line}"
            err.append(msg)
        print(err)
        folder = seleccionar_carpeta()
        msg_txt_err(folder, err)
        return False

def seleccionar_carpeta():
    """Abre un explorador de archivos para que el usuario seleccione una carpeta.
    Retorna la ruta de la carpeta seleccionada o None si el usuario cancela.
    """
    # Inicializar tkinter y ocultar la ventana principal
    ventana = tk.Tk()
    ventana.withdraw()  # Oculta la ventana principal
    # Configurar el ícono (usando archivo .ico o .png según tu sistema)
    try:
        ventana.iconbitmap(config('ICONO_VENTANA'))  # Reemplaza con la ruta de tu archivo .ico
    except:
        # Si el sistema no soporta iconbitmap (como en algunas distribuciones Linux), usa iconphoto con .png
        imagen = Image.open(config('PNG_VENTANA'))  # Reemplaza con la ruta de tu archivo .png
        icono = ImageTk.PhotoImage(imagen)
        ventana.iconphoto(True, icono)
    
    # Abrir el cuadro de diálogo para seleccionar la carpeta
    carpeta_seleccionada = filedialog.askdirectory(title="Seleccione una carpeta para guardar el archivo.")
    
    # Comprobar si la ventana se cerró sin seleccionar una carpeta
    if not carpeta_seleccionada:
        messagebox.showwarning("Advertencia", "No se seleccionó ninguna carpeta. Vuelva a iniciar el programa.")
        ventana.destroy()
        sys.exit()  # Finaliza la ejecución del programa
    
    # Cerrar la ventana de tkinter y retornar la carpeta seleccionada
    ventana.destroy()
    return carpeta_seleccionada


def guardar_xml(tree, path):
    """Guarda el XML en la ruta especificada con codificación y declaración."""
    with open(path, "wb") as file:
        tree.write(file, xml_declaration=True, encoding="UTF-8")

def eliminar_archivo(ruta):
    """Elimina un archivo en la ruta especificada si existe."""
    if os.path.isfile(ruta):
        os.remove(ruta)
        print(f"Archivo eliminado: {ruta}")
    else:
        print(f"El archivo no existe: {ruta}")


def msg_txt_err(filepath, mensaje):
    with open(f"{filepath}{config('ERR_XML_TXT_NAME')}", "a") as f:
        fecha_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f.write(f"->{fecha_hora}: {mensaje}\n")

def msg_txt(mensaje):
    with open(config('TXT_ERROR_PATH'), "a") as f:
        fecha_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        f.write(f"->{fecha_hora}:\n{mensaje}")

def msg_pass(msg, title = "Advertencia"):
    ventana = tk.Tk()
    ventana.withdraw()  # Oculta la ventana principal
    try:
        ventana.iconbitmap(config('ICONO_VENTANA'))  # Reemplaza con la ruta de tu archivo .ico
    except:
        # Si el sistema no soporta iconbitmap (como en algunas distribuciones Linux), usa iconphoto con .png
        imagen = Image.open(config('PNG_VENTANA'))  # Reemplaza con la ruta de tu archivo .png
        icono = ImageTk.PhotoImage(imagen)
        ventana.iconphoto(True, icono)
    messagebox.showwarning(title, msg)
    ventana.destroy()
    pass

def msg_exit(msg, title = "Error"):
    ventana = tk.Tk()
    ventana.withdraw()  # Oculta la ventana principal
    try:
        ventana.iconbitmap(config('ICONO_VENTANA'))  # Reemplaza con la ruta de tu archivo .ico
    except:
        # Si el sistema no soporta iconbitmap (como en algunas distribuciones Linux), usa iconphoto con .png
        imagen = Image.open(config('PNG_VENTANA'))  # Reemplaza con la ruta de tu archivo .png
        icono = ImageTk.PhotoImage(imagen)
        ventana.iconphoto(True, icono)
    messagebox.showwarning(title, msg)
    ventana.destroy()
    sys.exit()

def formatear_tabla(dataframe):
    # Determinar el ancho de cada columna en función del contenido
    col_widths = [max(len(str(cell)) for cell in dataframe[col].values) for col in dataframe.columns]
    col_widths = [max(len(str(col)), width) for col, width in zip(dataframe.columns, col_widths)]
    
    # Crear la línea de separación superior e inferior
    separator = "+" + "+".join("-" * (width + 2) for width in col_widths) + "+"

    # Formatear la cabecera de la tabla
    header = "|" + "|".join(f" {str(col).ljust(width)} " for col, width in zip(dataframe.columns, col_widths)) + "|"
    
    # Formatear cada fila
    rows = []
    for index, row in dataframe.iterrows():
        formatted_row = "|" + "|".join(f" {str(cell).ljust(width)} " for cell, width in zip(row, col_widths)) + "|"
        rows.append(formatted_row)

    # Combinar todas las partes de la tabla
    formatted_table = separator + "\n" + header + "\n" + separator + "\n" + "\n".join(rows) + "\n" + separator
    return formatted_table
