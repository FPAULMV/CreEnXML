import tkinter as tk 
from tkinter import scrolledtext, messagebox
import threading   
from tkinter import Menu
from PIL import Image, ImageTk
from pandastable import Table
from decouple import config  # Para leer archivos de configuración (como .env)
import pandas as pd
from lxml import etree as ET
import fns as f
from io import BytesIO
import os
from Crypto.Cipher import AES
import base64
from datetime import datetime
import requests
import json
import sys

class Principal:
    """
    Ejecucion del programa xml CRE
    """
    def __init__(self, root) -> None:
        self.root = root
        self.root.title("XML Creator")
        #self.root.resizable(False, False)
        root.iconbitmap(config('ICONO_VENTANA'))

        barra_menu = Menu(root)
        root.config(menu=barra_menu)
        menu_licencia = Menu(barra_menu, tearoff=0)
        barra_menu.add_cascade(label="Activar licencia", menu=menu_licencia)
        menu_licencia.add_command(label="Activar")

        """----------------------------------------------------------------------------------------------------"""
        # Frame de los botones.
        frame_botones = tk.Frame(root, border=1)
        # Ajustes
        x = 10
        y = 5
        f_n = "Consolas"
        f_z = 12
        f_s = "bold"
        b_fuente = (f_n,f_z)
        l_fuente = (f_n,10,f_s)

        config_b_subir = {
            "text":"Subir archivo",
            "font":b_fuente,
            "command":lambda: self.run_in_thread(self.subirarchivo),
            "width":14,
            "height":2
            #"bg":"",
            #"fg":"",
        }
        b_subir = tk.Button(frame_botones, **config_b_subir)

        config_b_crear = {
            "text":"Crear XML",
            "font":b_fuente,
            "command":lambda: self.run_in_thread(self.crearxml),
            "width":14,
            "height":2
            #"bg":"",
            #"fg":"",
        }
        b_crear = tk.Button(frame_botones, **config_b_crear)

        config_l_barrilescompras = {
            "text":"Barriles Comprados",
            "font":l_fuente
        }
        l_barrilescompras = tk.Label(frame_botones, **config_l_barrilescompras)

        config_l_barrilesventas = {
            "text":"Barriles Vendidos",
            "font":l_fuente
        }
        l_barrilesventas = tk.Label(frame_botones, **config_l_barrilesventas)

        config_t_barrilescompras = {
            "width":15,
            "height":1,
            "bg":"white",
            "fg":"black",
            "font":(f_n,15,f_s)
        }
        t_barriles_compras = tk.Text(frame_botones, **config_t_barrilescompras)
        t_barriles_compras.tag_configure("center", justify="center")
        t_barriles_compras.insert("1.0", "")  # Ejemplo de texto
        t_barriles_compras.tag_add("center", "1.0", "end")
        t_barriles_compras.config(state='disabled')


        config_t_barrilesventas = {
            "width":15,
            "height":1,
            "bg":"white",
            "fg":"black",
            "font":(f_n,15,f_s)
        }
        t_barriles_ventas = tk.Text(frame_botones, **config_t_barrilescompras)
        t_barriles_ventas.tag_configure("center", justify="center")
        t_barriles_ventas.insert("1.0", "")  # Ejemplo de texto
        t_barriles_ventas.tag_add("center", "1.0", "end")
        t_barriles_compras.config(state='disabled')

        # Posicion de los elementos.
        t_barriles_compras.grid(row=1, column=3, padx=x, pady=y)
        t_barriles_ventas.grid(row=1, column=4, padx=x, pady=y)
        l_barrilescompras.grid(row=0, column=3, padx=x, pady=y)
        l_barrilesventas.grid(row=0, column=4, padx=x, pady=y)
        b_subir.grid(row=0, rowspan=2, column=0, padx=x, pady=y)
        b_crear.grid(row=0, rowspan=2, column=1, padx=x, pady=y)
        frame_botones.grid(row=0, column=0, padx=x, pady=y)
        """----------------------------------------------------------------------------------------------------"""

        """----------------------------------------------------------------------------------------------------"""
        # Frame de la tabla.
        frame_tabla = tk.Frame(root, border=1)

        config_l_tabla = {
            "text":"Tabla",
            "font":l_fuente
        }
        l_tabla = tk.Label(frame_tabla, **config_l_tabla)

        config_tabla = {
            "wrap":tk.NONE,
            "width":120,
            "height":15,
            "font":(f_n, 8),
            "bg":"white",
            "fg":"black",
            "bd":1
        }
        self.label_tabla = tk.scrolledtext.ScrolledText(frame_tabla, **config_tabla)

        config_tablascrollbar_x = {
            "orient":"horizontal",
            "command":self.label_tabla.xview
        }
        tabla_scroll_x = tk.Scrollbar(frame_tabla, **config_tablascrollbar_x)

        # Posicion de los elementos.
        l_tabla.grid(row=0, column=0, sticky="W", padx=x, pady=y)
        self.label_tabla.grid(row=1, column=0, sticky="NSWE", padx=x, pady=y)
        tabla_scroll_x.grid(row=2, column=0, sticky="NSWE", padx=x, pady=y)
        frame_tabla.grid(row=1, column=0, padx=x, pady=y)

        # Frame de los mensajes
        frame_mensajes = tk.Frame(root)

        config_l_xml = {
            "text":"XML",
            "font":l_fuente
        }
        l_xml = tk.Label(frame_mensajes, **config_l_xml)

        config_l_msg = {
            "text":"Mensajes",
            "font":l_fuente
        }
        l_msg = tk.Label(frame_mensajes, **config_l_msg)

        config_xml = {
            "wrap":tk.NONE,
            "width":75,
            "height":15,
            "font":(f_n, 8),
            "bg":"white",
            "fg":"black",
            "bd":1
        }
        self.xml_text = tk.scrolledtext.ScrolledText(frame_mensajes, **config_xml)

        config_msg = {
            "wrap":tk.NONE,
            "width":35,
            "height":15,
            "font":(f_n, 8),
            "bg":"white",
            "fg":"black",
            "bd":1
        }
        self.msg_text = tk.scrolledtext.ScrolledText(frame_mensajes, **config_msg)
        self.msg_text.config(state='disabled')

        # Posicion de los elementos.
        l_xml.grid(row=0, column=0, sticky="W", padx=x, pady=y)
        l_msg.grid(row=0, column=1, sticky="W", padx=x, pady=y)
        self.xml_text.grid(row=1, column=0, padx=x, pady=y)
        self.msg_text.grid(row=1, column=1, padx=x, pady=y)
        frame_mensajes.grid(row=3, column=0, padx=x, pady=y)
        """----------------------------------------------------------------------------------------------------"""
        resultado = self.decrypt_data()
        fecha = pd.to_datetime(resultado['expiry_date']).date()
        """Agregar una seccion para validar la maquina contra la maquina de la licencia."""    
        date_now = datetime.now().date()
        print(fecha, date_now)
        if fecha >= date_now:
            print("Llave valida")
            pass
        else:
            print("Validando nueva llave")
            self.enviar_datos()
            print("Llave nueva cargada")

    def decrypt_data(self):
        try:
            with open(config('PATH_KEY'), 'r', encoding='utf-8') as archivo:
                self.contenido = archivo.read()

            # Configuración exacta como en PHP
            key = 'onlysecrets2024konlysecrets2024x'.encode('utf-8')
            iv = "1234567890123456".encode('utf-8')
            
            print("=== Iniciando desencriptación ===")
            print("Llave a desencriptar:", self.contenido)
            
            # Decodificar base64
            decoded = base64.b64decode(self.contenido)
            print("Longitud datos decodificados:", len(decoded))
            
            # Verificar IV
            received_iv = decoded[:16]
            print("IV recibido (hex):", received_iv.hex())
            print("IV esperado (hex):", iv.hex())
            
            # Obtener datos encriptados
            encrypted = decoded[16:]
            print("Longitud datos encriptados:", len(encrypted))
            print("Datos encriptados (hex):", encrypted.hex())
            
            # Crear cipher y desencriptar
            cipher = AES.new(key, AES.MODE_CBC, iv)
            decrypted_padded = cipher.decrypt(encrypted)
            
            # Remover padding PKCS7
            pad_length = decrypted_padded[-1]
            decrypted = decrypted_padded[:-pad_length].decode('utf-8')
            
            print("String desencriptado:", decrypted)
            
            # Separar las partes correctamente
            temp_parts = decrypted.split(':')
            
            # Reconstruir las partes
            parts = []
            # Los primeros 3 campos
            parts.extend(temp_parts[:3])
            # Fecha y hora (unir las partes 3,4,5)
            parts.append(':'.join(temp_parts[3:6]))
            # Estado (última parte)
            parts.append(temp_parts[-1])
            
            print("Partes:", parts)
            
            if len(parts) != 5:
                raise ValueError(f"Formato inválido: se encontraron {len(parts)} partes")
            
            # Reconstruir objeto
            self.result = {
                'product_id': parts[0],
                'client_id': parts[1],
                'machine_id': parts[2],
                'expiry_date': parts[3],
                'estado': parts[4]
            }
            return self.result
                
        except Exception as e:
            print(f"Error específico: {str(e)}")
            print("Tipo de error:", type(e))
            if 'decrypted_padded' in locals():
                print("Bytes completos (hex):", decrypted_padded.hex())
            raise e

    def enviar_datos(self):
        print(f"El contenido de la llave es: \n{self.contenido}")
        payload = {
            "license_data": self.contenido
        }
        
        url = "https://alvarium.mx/licverificationorg/validate.php"
        headers = {
            'Content-Type': 'application/json',
            'Accept': 'application/json',
            'Cache-Control': 'no-cache'
        }
        
        try:
            response = requests.post(
                url, 
                headers=headers, 
                data=json.dumps(payload)
            )
            if response.status_code == 200:
                response_data = response.json()
                print(response_data)
                print(response_data['success'])
                if response_data['success'] == False:
                    msg = f"Licencia invalida o inactiva."
                    print(msg)
                    f.msg_exit(msg)
                    sys.exit()
                elif response_data['success'] == True:
                    key = response_data['data']['hash']
                    print(f"La llave a validar es: \n{key}")
                    self.save_license_key(key)
                    return True
            else:
                msg = f"Request response <> 200"
                f.msg_exit(msg)
            
            
        except Exception as e:
            self.msg_text.insert(tk.END, f"Error al enviar datos: {str(e)}")

    def save_license_key(self, hash_value):
        """Guarda el hash en un archivo license.key"""
        try:
            # Si el archivo existe, lo borra
            if os.path.exists('license.key'):
                os.remove('license.key')
            
            # Crea el nuevo archivo y guarda el hash
            with open('license.key', 'w') as file:
                file.write(hash_value)
            
            messagebox.showinfo("Éxito", "Archivo license.key generado correctamente")
        except Exception as e:
            messagebox.showerror("Error", f"Error al generar archivo license.key: {str(e)}")

    df = None 
    xml_str = ""

    def subirarchivo(self):
        global df, xml_str
        f.pandas_consola()
        df = f.cargar_excel()

        """
        En esta sección, se transforman y ordenan los datos del Excel para ajustarlos al formato esperado en el XML.
        Las fechas inicial y final se definen según la fecha actual en el momento de ejecutar el programa. 
            Al ejecutarse, se espera que el archivo Excel contenga los datos correspondientes al periodo entre 
            el jueves más cercano anterior y el viernes previo a ese jueves. Si necesita reportar un periodo diferente, 
            se le solicitará que proporcione la fecha del viernes en que debió presentar su reporte ante la CRE.

        """
        # Combierte todos los datos a un objeto.
        df = df.astype(object)

        """USAR ESTA SECCION SOLO PARA PRUEBAS"""
        # fecha_prueba = '2024-11-07' # Opcional, hay que eliminarlo al crear el ejecutable. 
        # fechainicial, fechafinal = f.fechas_inicio_fin(fecha=fecha_prueba)
        # fechas = f.fechas_inicio_fin(fecha=fecha_prueba, valor=1)

        # Define la fecha que se esperaria pretende reportar, segun la fecha en la que se esta ejecutando el programa. 
        fechainicial, fechafinal = f.fechas_inicio_fin()
        fechas = f.fechas_inicio_fin(valor=1)


        # Obtener valores unicos de la fecha a reportar en Excel y los ordena de mas antiguo a mas reciente.
        df['Diaareportar'] = pd.to_datetime(df['Diaareportar'], errors='coerce')
        df_fechas = df['Diaareportar'] = df['Diaareportar'].dt.strftime('%Y-%m-%d')



        # Obtener conjunto unicos de ProductoId y SubProductoId y los ordena por 
        #   ProductoId de mayor a menor y SubProductoId de menor a mayor.
        productos = (df[['ProductoId', 'SubProductoId']].drop_duplicates())
        productos = productos.sort_values(by=['ProductoId', 'SubProductoId'], ascending=[False,True])

        """
        En esta seccion se valida que los datos formateados sean correctos.  
        """
        # Valida que las fechas del reporte coinicidan con las fechas que se 
        #   espera sean reportadas segun el dia que se ejecuta el programa.
        _f = set(fechas)
        _c = set(df_fechas)
        _n = []
        for i in _c:
            if i not in _f:
                _n.append(i)
        if _n:
            print(f"Las fechas a reportar\n{_n}. \nNo se encuentran en el rango de fechas que se deben de reportar")
        else:
            print(f"Todas las fechas son validas.")

        # Consulta a las api
        f.api_catalogoproducto()
        f.api_catalogosubproductos()

        # Elementos del reporte con los cuales se crea el XML. 
        fechasdelreporte = sorted(_c)

        # Crear la estructura XML
        root = ET.Element("ReporteVolumenes")
        permiso = ET.SubElement(root, 'Permiso', FechaFin=f"{fechafinal}", FechaInicio=f"{fechainicial}", Numero="H/1234/COM/1234", TipoReporte="1")
        for fecha in fechasdelreporte:
            fecha_elemen = ET.SubElement(permiso, 'Fecha', Diaareportar=f"{fecha}")
            df_fil = df[df['Diaareportar'] == fecha]
            df_fil['ProductoId'] = df_fil['ProductoId'].astype(str)
            df_fil['SubProductoId'] = df_fil['SubProductoId'].astype(str)
            df_fil['TipoMov'] = df_fil['TipoMov'].astype(str).str.lower()
            df_prod = df_fil[['ProductoId', 'SubProductoId']].drop_duplicates()
            df_prod = df_prod.sort_values(by=['ProductoId', 'SubProductoId'], ascending=[False, True])
            
            for _, row in df_prod.iterrows():
                p = row['ProductoId']
                sp = row['SubProductoId']
                producto = ET.SubElement(fecha_elemen, 'Producto', ProductoId=f"{p}", SubProductoId=f"{sp}")
                ventasnacional = ET.SubElement(producto, 'VentasNacional')
                # Filtrar el DataFrame para obtener solo las ventas del Producto y SubProducto actual
                df_ventas = df_fil[(df_fil['ProductoId'] == str(p)) & (df_fil['SubProductoId'] == str(sp)) & (df_fil['TipoMov'] == 'venta')]
                df_ventas['VolumenVendido'] = df_ventas['VolumenVendido'].apply(lambda x: f"{float(x):.2f}" if x != '' else '')
                df_ventas['Entidad'] = df_ventas['Entidad'].apply(lambda x: str(int(float(x))) if x != '' and pd.notna(x) else '')
                df_ventas['Municipio'] = df_ventas['Municipio'].apply(lambda x: str(int(float(x))) if x != '' and pd.notna(x) else '')
                df_ventas = df_ventas.fillna("")
                # Crear elementos PermisionarioCRECliente para cada fila filtrada en df_ventas
                for _, venta_row in df_ventas.iterrows():
                    if venta_row['Entidad'] == "":
                        if venta_row['CostoFlete'] == "":
                            permisionario = ET.SubElement(ventasnacional, 'PermisionarioCRECliente', NumeroPermisoCRECliente=f"{venta_row['NumeroPermisoCRECliente']}",PrecioVenta=f"{venta_row['PrecioVenta']}",VolumenVendido=f"{venta_row['VolumenVendido']}")
                            permisionario.text = ''
                        else:
                            permisionario = ET.SubElement(ventasnacional, 'PermisionarioCRECliente', NumeroPermisoCRECliente=f"{venta_row['NumeroPermisoCRECliente']}",PrecioVenta=f"{venta_row['PrecioVenta']}",VolumenVendido=f"{venta_row['VolumenVendido']}")
                            flete = ET.SubElement(permisionario, 'FleteVNPermisionarioCRE', CostoFlete=f"{venta_row['CostoFlete']}", PermisoTransportista=f"{venta_row['PermisoTransportista']}")
                        permisionario.text = '' 
                    else:
                        if venta_row['CostoFlete'] == "":
                            permisionario = ET.SubElement(ventasnacional, 'PermisionarioCRECliente', NumeroPermisoCRECliente=f"{venta_row['NumeroPermisoCRECliente']}",PrecioVenta=f"{venta_row['PrecioVenta']}",VolumenVendido=f"{venta_row['VolumenVendido']}", Entidad=f"{venta_row['Entidad']}", Municipio=f"{venta_row['Municipio']}")
                            permisionario.text = ''
                        else:
                            permisionario = ET.SubElement(ventasnacional, 'PermisionarioCRECliente', NumeroPermisoCRECliente=f"{venta_row['NumeroPermisoCRECliente']}",PrecioVenta=f"{venta_row['PrecioVenta']}",VolumenVendido=f"{venta_row['VolumenVendido']}", Entidad=f"{venta_row['Entidad']}", Municipio=f"{venta_row['Municipio']}")
                            flete = ET.SubElement(permisionario, 'FleteVNPermisionarioCRE', CostoFlete=f"{venta_row['CostoFlete']}", PermisoTransportista=f"{venta_row['PermisoTransportista']}")
                        permisionario.text = '' 
                ventasnacional.text = ''
                comprasnacional = ET.SubElement(producto, 'ComprasNacional')
                df_compras = df_fil[(df_fil['ProductoId'] == str(p)) & (df_fil['SubProductoId'] == str(sp)) & (df_fil['TipoMov'] == 'compra')]
                df_compras['VolumenComprado'] = df_compras['VolumenComprado'].apply(lambda x: f"{float(x):.2f}" if x != '' else '')
                df_compras['CostoFlete'] = df_compras['CostoFlete'].apply(lambda x: f"{float(x):.2f}" if x != '' else '')
                df_compras['TipoDescuentoId'] = df_compras['TipoDescuentoId'].astype(int)
                df_compras['CostoFlete'] = df_compras['CostoFlete'].replace("nan","")
                df_compras = df_compras.fillna("")
                for _, compra_row in df_compras.iterrows():
                    if compra_row['CostoFlete'] == "":
                        permisionariocompra = ET.SubElement(comprasnacional, 'PermisionarioCREProveedor', NumeroPermisoCREProveedor=f"{compra_row['NumeroPermisoCREProveedor']}", PrecioCompra=f"{compra_row['PrecioCompra']}", VolumenComprado=f"{compra_row['VolumenComprado']}")
                        descuento = ET.SubElement(permisionariocompra, 'DescuentoCompraNacionalPermisionarioCRE', PrecioDescuentoIncluido=f"{compra_row['PrecioDescuentoIncluido']}", TipoDescuentoId=f"{compra_row['TipoDescuentoId']}")
                        permisionariocompra.text = ''
                    else:
                        permisionariocompra = ET.SubElement(comprasnacional, 'PermisionarioCREProveedor', NumeroPermisoCREProveedor=f"{compra_row['NumeroPermisoCREProveedor']}", PrecioCompra=f"{compra_row['PrecioCompra']}", VolumenComprado=f"{compra_row['VolumenComprado']}")
                        flete = ET.SubElement(permisionariocompra, 'FleteCompraNacionalPermisionarioCRE', CostoFlete=f"{compra_row['CostoFlete']}", PermisoTransportista=f"{compra_row['PermisoTransportista']}")
                        descuento = ET.SubElement(permisionariocompra, 'DescuentoCompraNacionalPermisionarioCRE', PrecioDescuentoIncluido=f"{compra_row['PrecioDescuentoIncluido']}", TipoDescuentoId=f"{compra_row['TipoDescuentoId']}")
                        permisionariocompra.text = ''
                comprasnacional.text = ""
                producto.text = ''
            fecha_elemen.text = ''

        # Identacion automatica 
        ET.indent(root)
        # Crea la raíz
        tree = ET.ElementTree(root)

        xml_buffer = BytesIO()
        tree = ET.ElementTree(root)
        ET.indent(root)
        tree.write(xml_buffer, encoding='utf-8', xml_declaration=True)
        
        # Convertir el XML a una cadena para su retorno
        xml_str = xml_buffer.getvalue().decode('utf-8')

        # Validar el XML con el XSD (si corresponde)
        if f.validar_xml_con_xsd(xml_str, config('XSD_VALIDATE')):
            print("XML válido")
            dataframe = f.formatear_tabla(df)
            self.print_text(self.label_tabla, dataframe)
            self.print_text(self.xml_text, xml_str)
            return df, xml_str
        else:
            print("XML no válido")
        
        # Retornar el DataFrame y la cadena XML
        
    def crearxml(self):
        global xml_str
        # Seleccionar carpeta de destino
        folder = f.seleccionar_carpeta()
        final_path = os.path.join(folder, f"{config('OUTPUT_XML_NAME')}.xml")
        
        # Guardar el contenido de xml_str en el archivo XML
        with open(final_path, 'w', encoding='utf-8') as file:
            file.write(xml_str)
        
        # Notificación de éxito
        title = "Finalizado"
        msg = f"Archivo XML guardado en:\n{final_path}"
        print(msg)
        f.msg_pass(msg, title)
        
    def print_text(self, element, message: str):
        self.element = element
        """
        La función solicita como parámetro un mensaje en texto el cual 
        imprimirá en la pantalla tipo CMD.

        - Cambia el estado de la pantalla a normal, lo que permite ingresar texto.
        - Agrega el mensaje a la pantalla.
        - Mueve la vista al final de la pantalla.
        - Bloquea nuevamente la pantalla. 
        """
        self.element.config(state="normal")
        self.element.delete("1.0", tk.END)
        self.element.insert(tk.END, str(message) +'\n')
        self.element.config(state="disabled")

    # Función que ejecuta las tareas en un hilo separado

    def run_in_thread(self, target, *args):
        """
        Requiere como argumentos un objeto que llamara para ejecutarse en segundo plano. 
        Seguido de los argumentos que se quieren ejecutar en este objeto. 
        """
        thread = threading.Thread(target=target, args=args)
        thread.start() 

if __name__ == '__main__':
    print(f"Ejecutando módulo {__name__}")
    root = tk.Tk()
    app = Principal(root)
    root.mainloop()