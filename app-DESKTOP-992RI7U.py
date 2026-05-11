import flet as ft
import webbrowser
import tkinter as tk
import re
from tkinter import filedialog
import os
import sys
from datetime import datetime
import win32com.client

from backend.excel_reader import leer_excel_multiple
from backend.excel_generator import llenar_excel
from backend.pdf_generator import excel_a_pdf
from backend.xml_generator import generar_xml  

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
    page.scroll = None
    page.title = "Generador de Documentos"
    page.theme_mode = ft.ThemeMode.DARK
    page.window_width = 700
    page.window_height = 850
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

        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )

        if file_path:
            ruta_archivo["path"] = file_path
            archivo_cargado.value = "Archivo cargado ✔"
            archivo_cargado.color = "green"
            estado.value = "Excel listo"
            page.update()

    # -------------------------
    # 📁 SELECTOR CARPETA (FLET)
    # -------------------------


    def seleccionar_carpeta(e):
        root = tk.Tk()
        root.withdraw()

        carpeta = filedialog.askdirectory()

        if carpeta:
            carpeta_salida["path"] = carpeta
            estado.value = "📁 Carpeta seleccionada"
            estado.color = "green"
            page.update()

    # -------------------------
    # 📁 CREAR CARPETA POR BL
    # -------------------------
    def crear_carpeta_bl(base_path, bl):
        nombre = limpiar_nombre_archivo(bl)
        ruta = os.path.join(base_path, nombre)
        os.makedirs(ruta, exist_ok=True)
        return ruta

    # -------------------------
    # ⚙️ GENERAR DOCUMENTOS
    # -------------------------
    def generar(e):
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

            lista = leer_excel_multiple(ruta_archivo["path"])
            total = len(lista)

            #agregar_log(f"📦 Total BL: {total}")

            docx_pdf_lista = []

            for i, datos in enumerate(lista, start=1):

                if cancelar["activo"]:
                    estado.value = "Proceso cancelado"
                    estado.color = "red"
                    break

                bl = datos["bl"]
                #agregar_log(f"▶ Procesando {bl}")

                try:
                    carpeta_bl = crear_carpeta_bl(carpeta_salida["path"], bl)
                    nombre = limpiar_nombre_archivo(bl)

                    excel = os.path.join(carpeta_bl, f"{nombre}.xlsx")
                    pdf = os.path.join(carpeta_bl, f"{nombre}.pdf")
                    xml = os.path.join(carpeta_bl, f"{nombre}.xml")

                    
                    ruta_xml_template = ruta_recursos("plantillas/plantilla.xml")
                    ruta_excel_template = ruta_recursos("plantillas/template.xlsx")

                    generar_xml(ruta_xml_template, xml, datos)
                    llenar_excel(ruta_excel_template, excel, datos)
                    excel_a_pdf(excel, pdf)
                    #agregar_log(f"✔ OK {bl}")

                except Exception as err_bl:
                    print(f"ERROR EN {bl}: {err_bl}")  # 👈 IMPORTANTE
                    estado.value = f"Error en {bl}"
                    estado.color = "red"

                progreso.value = i / total
                texto_progreso.value = f"Procesando {i} de {total}"
                page.update()

            loader.visible = True

            if not cancelar["activo"]:
                estado.value = "✔ Proceso completado"
                texto_progreso.value = "Completado ✔"
                estado.color = "green"

        except Exception as err:
            estado.value = str(err)
            estado.color = "red"

        page.update()    

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

                ft.Container(height=5),

                ft.ElevatedButton(
                    "Generar PDF + XML",
                    on_click=generar,
                    width=280,
                    bgcolor="#2e7d32",   # verde empresa
                    color="white"
                ),

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

                estado,

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
        width=360,

        # 🔥 SOMBRA SUTIL
        shadow=ft.BoxShadow(
            blur_radius=20,
            color="#000000",
            offset=ft.Offset(2, 4)
        )
    )

    page.add(card)


# -------------------------
# ▶️ RUN
# -------------------------
ft.app(target=main)