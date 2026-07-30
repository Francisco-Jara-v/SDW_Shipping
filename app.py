import asyncio

import flet as ft
import webbrowser
import tkinter as tk
import re
from tkinter import filedialog
import os
import sys
import tempfile
import subprocess
import threading
import requests
from datetime import datetime
import win32com.client
from update_checker import hay_actualizacion
from version import VERSION
import webbrowser
from urllib.parse import urlparse
import ctypes
import time

from backend.marca_agua import agregar_marca_agua
from backend.excel_reader import leer_excel_multiple
from backend.excel_generator import llenar_excel
from backend.pdf_generator import excel_a_pdf
from backend.xml_generator import generar_xml  


data_bl = []
tabla_rows = []

seleccionados_count = ft.Text("Seleccionados: 0")
busqueda = ft.TextField(
    hint_text="Buscar BL o Consignee...",
    width=300,
)
lista_filtrada = []

def ruta_recursos(rel_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, rel_path)


# -------------------------
# 🔧 UTILIDAD
# -------------------------
def limpiar_nombre_archivo(texto):
    if not texto:
        return "SIN_NOMBRE"

    texto = texto.strip()
    texto = re.sub(r'[\\/*?:"<>|]', '', texto)
    texto = texto.replace("\n", " ")
    texto = "_".join(texto.split())

    return texto




# -------------------------
# 🚀 APP
# -------------------------
def main(page: ft.Page):
    # CONFIG
    page.window_resizable = True
    page.window.icon = ruta_recursos("plantillas/logo.ico")
    page.scroll = ft.ScrollMode.AUTO
    page.title = "Generador de Documentos"
    page.theme_mode = ft.ThemeMode.DARK
    page.window_width = 1000
    page.window_height = 1500
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    #page.bgcolor = "#121212"

    progreso = ft.ProgressBar(width=300, value=0)
    loader = ft.ProgressRing(visible = False)
    texto_progreso = ft.Text("")
    #log_output = ft.Text(value="", selectable=True)
    cancelar={"activo":False}
    ruta_archivo = {"path": None}
    carpeta_salida = {"path": ""}

    def revisar_actualizaciones(page: ft.Page):
        """Ejecuta la comprobación y muestra el diálogo directamente en el overlay de la página."""

        async def _tarea_verificacion():
            try:
                print(f"[DEBUG] Comprobando actualización. Versión local actual: {VERSION}")

                # Consultamos GitHub en segundo plano sin congelar la interfaz
                hay, release = await asyncio.to_thread(hay_actualizacion)
                print(f"[DEBUG] Resultado de hay_actualizacion(): hay={hay}")

                if not hay or not release:
                    print("[DEBUG] No se requiere actualización.")
                    return

                # Elementos de control de descarga
                progress_bar = ft.ProgressBar(visible=False, value=0)
                status_text = ft.Text("", size=12, color=ft.Colors.GREY_400)

                def cerrar_dialogo(e):
                    dialog.open = False
                    page.update()

                def iniciar_actualizacion_automatica(e):
                    btn_descargar.disabled = True
                    btn_mas_tarde.disabled = True
                    progress_bar.visible = True
                    status_text.value = "Conectando al servidor..."
                    page.update()

                    async def _tarea_descarga():
                        try:
                            url = release.get("download_url")

                            if not url:
                                status_text.value = "Error: No se encontró la URL de descarga."
                                btn_mas_tarde.disabled = False
                                page.update()
                                return

                            # 1. Extracción limpia de la extensión ignorando parámetros como '?token=...'
                            parsed_url = urlparse(url)
                            path_url = parsed_url.path.lower()
                            file_extension = ".exe" if path_url.endswith(".exe") else (".zip" if path_url.endswith(".zip") else ".exe")

                            temp_dir = tempfile.gettempdir()
                            version_release = release.get('version', '2.0.0')
                            installer_path = os.path.join(temp_dir, f"SDW_Setup_{version_release}{file_extension}")

                            print(f"[DEBUG] URL detectada: {url}")
                            print(f"[DEBUG] Guardando en: {installer_path}")

                            def _descargar_archivo():
                                response = requests.get(url, stream=True, timeout=30)
                                response.raise_for_status()

                                total_bytes = int(response.headers.get("content-length", 0))
                                downloaded = 0
                                last_update_time = time.time()
                                last_percentage = -1.0

                                with open(installer_path, "wb") as file:
                                    for chunk in response.iter_content(chunk_size=65536):
                                        if not chunk:
                                            continue
                                        file.write(chunk)
                                        downloaded += len(chunk)

                                        if total_bytes > 0:
                                            porcentaje = downloaded / total_bytes
                                            pct_100 = round(porcentaje * 100, 1)
                                            now = time.time()

                                            # Throttling: actualiza UI máximo cada 100ms para evitar lag
                                            if (now - last_update_time > 0.1 and pct_100 != last_percentage) or downloaded == total_bytes:
                                                last_update_time = now
                                                last_percentage = pct_100
                                                progress_bar.value = porcentaje
                                                status_text.value = f"Descargando: {pct_100:.1f}%"
                                                page.update()

                            # Descarga asíncrona en hilo secundario
                            await asyncio.to_thread(_descargar_archivo)

                            status_text.value = "Ejecutando instalador..."
                            page.update()

                            # 2. Ejecución con elevación de administrador y cierre limpio de app
                            if installer_path.lower().endswith(".exe"):
                                await asyncio.sleep(0.5)
                                args = "/VERYSILENT /SUPPRESSMSGBOXES /CLOSEAPPLICATIONS"

                                if os.name == "nt":
                                    print(f"[DEBUG] Solicitando elevación UAC para: {installer_path}")
                                    resultado = ctypes.windll.shell32.ShellExecuteW(
                                        None,
                                        "runas",          # Muestra la ventana UAC de Windows
                                        installer_path,
                                        args,             # Parámetros silenciosos de Inno Setup
                                        None,
                                        1
                                    )

                                    if resultado > 32:
                                        print("[DEBUG] Instalador lanzado. Cerrando aplicación actual...")
                                        try:
                                            await page.window.close()
                                        except Exception:
                                            pass
                                        os._exit(0)  # Cierra inmediatamente sin esperar a que se limpien los hilos
                                    else:
                                        print(f"[WARN] ShellExecuteW código {resultado}. Probando fallback...")
                                        subprocess.Popen(f'"{installer_path}" {args}', shell=True)
                                        os._exit(0)
                                else:
                                    subprocess.Popen([installer_path])
                                    os._exit(0)
                            else:
                                status_text.value = f"Descargado en: {installer_path}"
                                btn_mas_tarde.disabled = False
                                page.update()

                        except Exception as ex:
                            print(f"[ERROR DESCARGA] {ex}")
                            status_text.value = f"Error en descarga: {ex}"
                            btn_mas_tarde.disabled = False
                            page.update()

                    page.run_task(_tarea_descarga)

                btn_mas_tarde = ft.TextButton("Más tarde", on_click=cerrar_dialogo)
                btn_descargar = ft.Button("Actualizar Ahora", on_click=iniciar_actualizacion_automatica)

                # Construcción del diálogo
                dialog = ft.AlertDialog(
                    modal=True,
                    title=ft.Text("🚀 Nueva actualización disponible"),
                    content=ft.Column(
                        [
                            ft.Text(f"Versión instalada: v{VERSION}"),
                            ft.Text(f"Nueva versión disponible: v{release.get('version', 'S/N')}"),
                            ft.Divider(),
                            ft.Text("Novedades:"),
                            ft.Text(
                                release.get("descripcion", "Sin detalles"),
                                size=12,
                                selectable=True
                            ),
                            status_text,
                            progress_bar,
                        ],
                        tight=True,
                        scroll=ft.ScrollMode.AUTO,
                        height=200,
                        width=400,
                    ),
                    actions=[btn_mas_tarde, btn_descargar]
                )

                print("[DEBUG] Agregando diálogo a page.overlay...")
                page.overlay.append(dialog)
                dialog.open = True
                page.update()
                print("[DEBUG] ¡Diálogo enviado a pantalla con éxito!")

            except Exception as err:
                print(f"[ERROR EN CHECKER DE ACTUALIZACIONES] {err}")

        page.run_task(_tarea_verificacion)


    # -------------------------
    # 🎨 COLORES
    # -------------------------
    def color_card():
        return "#1e1e1e" if page.theme_mode == ft.ThemeMode.DARK else "#ffffff"

    def color_texto_secundario():
        return "grey" if page.theme_mode == ft.ThemeMode.DARK else "black"

    def icono_tema():
        return (
            ft.Icons.WB_SUNNY
            if page.theme_mode == ft.ThemeMode.DARK
            else ft.Icons.NIGHTLIGHT
        )

    # -------------------------
    # 🧾 UI TEXTOS
    # -------------------------
    archivo_cargado = ft.Text("Ningún archivo cargado", color=color_texto_secundario())
    estado = ft.Text(
        "Esperando acción...",
        size=13,
        weight="w500"
    )

    logo = ft.Image(
        src=ruta_recursos("plantillas/LOGO.png"),
        width=120,
        height=120,
        fit="contain"
    )

    # -------------------------
    # 📂 SELECTOR EXCEL (TKINTER)
    # -------------------------
    def seleccionar_archivo(e):
        root = tk.Tk()
        root.withdraw()
        
        root.attributes("-topmost", True)  # Asegura que el diálogo esté al frente
        root.update()  # Actualiza la ventana para aplicar el cambio
    
        file_path = filedialog.askopenfilename(
            parent=root,
            title="Selecciona el archivo Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        root.destroy()  # Cierra la ventana de Tkinter después de seleccionar el archivo

        if file_path:
            ruta_archivo["path"] = file_path
            archivo_cargado.value = "Archivo cargado ✔"
            archivo_cargado.color = "green"
            estado.value = "Excel listo"
            cargar_tabla(file_path)
            page.update()

    def cargar_tabla(ruta):
        global data_bl, tabla_rows

        data_bl = leer_excel_multiple(ruta)
        tabla_rows.clear()
        contenedor_filas.controls.clear()

        for d in data_bl:
            checkbox = ft.Checkbox(
                value=True,
                on_change=lambda e: actualizar_contador()
            )

            row = ft.Row(
                [
                    checkbox,
                    ft.Text(d["bl"], width=180),
                    ft.Text(d["consignee"], width=250),
                ]
            )

            row.data = checkbox
            tabla_rows.append(row)
            contenedor_filas.controls.append(row)

        actualizar_contador()
        page.update()
        
    def actualizar_contador():
        count = sum(1 for r in tabla_rows if r.data.value)
        seleccionados_count.value = f"Seleccionados: {count}"
        page.update()
    
    
    
        
    def filtrar_tabla(e=None):
        texto = (busqueda.value or "").lower()

        contenedor_filas.controls.clear()

        count = 0

        for i, d in enumerate(data_bl):
            if texto in d["bl"].lower() or texto in d["consignee"].lower():
                row = tabla_rows[i]
                contenedor_filas.controls.append(row)

                if row.data.value:
                    count += 1

        seleccionados_count.value = f"Seleccionados: {count}"
        page.update()

    busqueda.on_change = filtrar_tabla

    # -------------------------
    # 📁 SELECTOR CARPETA (FLET)
    # -------------------------


    def seleccionar_carpeta(e):
        root = tk.Tk()
        root.withdraw()
        
        root.attributes("-topmost", True)  # Asegura que el diálogo esté al frente
        root.update()  # Actualiza la ventana para aplicar el cambio

        carpeta = filedialog.askdirectory()
        
        root.destroy()  # Cierra la ventana de Tkinter después de seleccionar la carpeta

        if carpeta:
            carpeta_salida["path"] = carpeta
            estado.value = "📁 Carpeta seleccionada"
            estado.color = "green"
            page.update()

    # -------------------------
    # 📁 CREAR CARPETA POR BL
    # -------------------------
    #def crear_carpeta_bl(base_path, bl):
    #    nombre = limpiar_nombre_archivo(bl)
    #    ruta = os.path.join(base_path, nombre)
    #    os.makedirs(ruta, exist_ok=True)
    #    return ruta

    # -------------------------
    # ⚙️ GENERAR DOCUMENTOS
    # -------------------------
    async def generar(e):
        cancelar["activo"] = False
        #log_output.value = ""

        if not ruta_archivo["path"]:
            estado.value = "Selecciona Excel"
            return

        if not carpeta_salida["path"]:
            estado.value = "Selecciona carpeta"
            return

        try:
            loader.visible = True
            progreso.value = 0
            page.update()

            lista = []

            for i, row in enumerate(tabla_rows):
                if row.data.value:
                    lista.append(data_bl[i])

            total = len(lista)

            #agregar_log(f"📦 Total BL: {total}")

            docx_pdf_lista = []
            import asyncio
            
            for i, datos in enumerate(lista, start=1):
                bl = datos["bl"]
                
                if cancelar["activo"]:
                    estado.value = f"Proceso cancelado en {bl}"
                    estado.color = "red"
                    break

                
                
                #agregar_log(f"▶ Procesando {bl}")

                try:
                    
                    nombre = limpiar_nombre_archivo(bl)

                    excel = os.path.join(carpeta_salida["path"], f"{nombre}.xlsx")
                    pdf = os.path.join(carpeta_salida["path"], f"original_{nombre}.pdf")
                    pdf_copia = os.path.join(
                                            carpeta_salida["path"],
                                            f"COPIA_NO_NEGOCIABLE_{nombre}.pdf"
                                        )
                    xml = os.path.join(carpeta_salida["path"], f"{nombre}.xml")
                    
                    
                    ruta_xml_template = ruta_recursos("plantillas/plantilla.xml")
                    ruta_excel_template = ruta_recursos("plantillas/template.xlsx")
                    
                    estado.value = f"Generando XML: {bl}"
                    page.update()
                    
                    await asyncio.to_thread(generar_xml, ruta_xml_template, xml, datos)
                    
                    if cancelar["activo"]:
                        estado.value = f"Cancelado en {bl}"
                        estado.color = "red"
                        break
                    
                    estado.value = f"Generando Excel: {bl}"
                    page.update()
                    await asyncio.to_thread(llenar_excel, ruta_excel_template, excel, datos)
                    if cancelar["activo"]:
                        estado.value = f"Cancelado en {bl}"
                        estado.color = "red"
                        break
                    estado.value = f"Generando PDF: {bl}"
                    page.update()
                    await asyncio.to_thread(
                        excel_a_pdf,
                        excel,
                        pdf,
                        1
                    )
                    
                    # PDF Copia No Negociable (Hoja 2)
                    await asyncio.to_thread(
                        excel_a_pdf,
                        excel,
                        pdf_copia,
                        2
                    )
                    if cancelar["activo"]:
                        estado.value = f"Cancelado en {bl}"
                        estado.color = "red"
                        break
                    

                    estado.value = f"Aplicando marca de agua: {bl}"
                    page.update()
                    pdf_final = pdf.replace(".pdf", "_final.pdf")

                    await asyncio.to_thread(
                        agregar_marca_agua,
                        pdf,
                        pdf_final,
                        ruta_recursos("plantillas/marca_agua.png")
                    )
                    
                    import time
                    time.sleep(0.5)  # Asegura que el archivo esté listo antes de moverlo

                    if os.path.exists(pdf_final):
                        os.remove(pdf)
                        os.rename(pdf_final, pdf)
                    else:
                        raise Exception("Error al aplicar marca de agua")
                    #agregar_log(f"✔ OK {bl}")

                except Exception as err_bl:
                    print(f"ERROR EN {bl}: {err_bl}")  # 👈 IMPORTANTE
                    estado.value = f"Error en {bl}"
                    estado.color = "red"

                progreso.value = i / total
                texto_progreso.value = f"Procesando {i} de {total}"
                page.update()  # Actualización asíncrona para mejor rendimiento
                await asyncio.sleep(0.01)  # Permite que la UI se refresque
                

            # Convertir PDFs SOLO si no canceló
            if not cancelar["activo"]:
                from docx2pdf import convert
                convert(carpeta_salida["path"])
                #agregar_log("📄 PDFs generados")
        
        finally:
            loader.visible = False
            if not cancelar["activo"]:
                estado.value = "✔ Proceso completado"
                texto_progreso.value = "Completado ✔"
                estado.color = "green"
            try:
                page.update()
            except:
                pass
    
    def seleccionar_todos():
        for row in tabla_rows:
            row.data.value = True

        actualizar_contador()
        filtrar_tabla()
        
    def invertir_seleccion():
        for row in tabla_rows:
            row.data.value = not row.data.value

        actualizar_contador()
        filtrar_tabla()
    
    #def agregar_log(msg):
    #    log_output.value += msg + "\n"
    #    page.update()
    
    def cancelar_proceso(e):
        cancelar["activo"] = True
        #agregar_log("⛔ Proceso cancelado por el usuario")
        
    # -------------------------
    # 📧 CONTACTO
    # -------------------------
    def contacto(e):
        webbrowser.open("mailto:francisco.jara.valdes95@gmail.com")

    # -------------------------
    # 🌙 CAMBIAR TEMA
    # -------------------------
    def cambiar_tema(e):
        page.theme_mode = (
            ft.ThemeMode.LIGHT
            if page.theme_mode == ft.ThemeMode.DARK
            else ft.ThemeMode.DARK
        )

        card.bgcolor = color_card()
        archivo_cargado.color = color_texto_secundario()
        boton_tema.icon = icono_tema()

        page.update()

    boton_tema = ft.IconButton(
        icon=icono_tema(),
        on_click=cambiar_tema
    )

    # -------------------------
    # 🧩 CARD UI
    # -------------------------
    
    header = ft.Row(
        [
            ft.Text("✔", width=40),
            ft.Text("BL", width=180, weight="bold"),
            ft.Text("Consignee", width=250, weight="bold"),
        ]
    )

    contenedor_filas = ft.Column(
        scroll=ft.ScrollMode.AUTO,
        height=250
    )
    
    card = ft.Container(
        content=ft.Column(
            [
                logo,

                ft.Text(
                    "Generador de Documentos",
                    size=20,
                    weight="bold"
                ),

                boton_tema,

                ft.Divider(height=20),

                ft.ElevatedButton(
                    "Seleccionar Excel",
                    on_click=seleccionar_archivo,
                    width=280
                ),

                archivo_cargado,

                ft.ElevatedButton(
                    "Seleccionar carpeta destino",
                    on_click=seleccionar_carpeta,
                    width=280
                ),
                
                ft.Column(
                    [
                        busqueda,
                    
                        seleccionados_count,

                        ft.Row(
                            [
                                ft.ElevatedButton("Seleccionar todos", on_click=lambda e: seleccionar_todos()),
                                ft.ElevatedButton("Invertir selección", on_click=lambda e: invertir_seleccion()),
                            ],
                            alignment=ft.MainAxisAlignment.CENTER
                        ),

                        header,

                        ft.Container(
                            content=contenedor_filas,
                            height=250
                        )
                    ]
                ),
                
                

                ft.Container(height=5),

                
                #ft.ElevatedButton(
                #    "Seleccionar todos",
                #    on_click=lambda e: seleccionar_todos(),
                #    width=200
                #),
                
                
                ft.ElevatedButton(
                    "Generar PDF + XML",
                    on_click=lambda e: page.run_task(generar, e),
                    width=280,
                    bgcolor="#2e7d32",   # verde empresa
                    color="white"
                ),
                
                estado,
                

                ft.Container(height=10),

                loader,
                progreso,
                texto_progreso,

                ft.Container(height=15),

                ft.ElevatedButton(
                    "Cancelar proceso",
                    on_click=cancelar_proceso,
                    width=280,
                    bgcolor="#c62828",
                    color="white"
                ),

                ft.Container(height=15),

                

                ft.Divider(height=20),

                ft.Text("Contacto", size=12, color="grey"),

                ft.ElevatedButton(
                    "Enviar correo",
                    on_click=contacto,
                    width=200
                )
            ],
            
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=10
        ),
        padding=25,
        border_radius=15,
        bgcolor=color_card(),
        expand=True,
        
        # 🔥 SOMBRA SUTIL
        shadow=ft.BoxShadow(
            blur_radius=20,
            color="#000000",
            offset=ft.Offset(2, 4)
        )
    )

    page.add(
        ft.Container(
            content=card,
            padding=20,
            expand=True
        )
    )
    
    revisar_actualizaciones(page)


# -------------------------
# ▶️ RUN
# -------------------------
ft.run(main)