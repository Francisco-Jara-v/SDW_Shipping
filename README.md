# SDW Shipping - Document Generator & Data Processor 🚢

Sistema profesional de automatización logística diseñado para transformar datos brutos de planillas Excel en documentos operativos críticos (PDF y XML). Este proyecto marca la evolución de un sistema funcional hacia un producto sólido, escalable y listo para entornos empresariales.

## 🚀 Características Principales

### 📄 Generación de Documentos PDF
* **Motor de Capas Inteligente:** Implementación de marcas de agua dinámicas situadas por debajo de la capa de texto, manteniendo la legibilidad y la capacidad de búsqueda/indexación.
* **Ajuste de Filas Adaptativo:** Lógica personalizada para el cálculo de altura de filas en celdas combinadas, evitando cortes de texto en descripciones largas.
* **Numeración Automática:** Generación dinámica de ítems (ej. `1. TRUCK HEAD`, `2. CHASIS`) directamente desde el flujo de datos.

### 🧬 Procesamiento de Datos Estructurados
* **Limpieza de Datos Robusta:** Filtros avanzados para la conversión de tipos (pesos, fletes, fechas) que corrigen inconsistencias comunes en formatos de celdas de Excel.
* **Generador XML:** Creación de manifiestos electrónicos mediante plantillas, permitiendo la integración con sistemas de aduanas y tracking.
* **Mapeo Inteligente:** Capacidad de agrupar múltiples especificaciones por número de BL, consolidando la información de transporte.

### 📦 Distribución Profesional
* **Instalador EXE:** Empaquetado con PyInstaller e Inno Setup, permitiendo que el cliente final use la herramienta sin necesidad de instalar Python o dependencias.
* **Interfaz de Usuario:** Basada en el framework Flet para una experiencia de escritorio fluida y moderna.

## 🛠️ Stack Tecnológico

* **Lenguaje:** Python 3.10+
* **Manipulación de Datos:** Pandas, OpenPyXl.
* **Procesamiento de Archivos:** PyMuPDF (fitz), PyPDF2.
* **Interfaz Gráfica:** Flet (Flutter para Python).
* **Compilación:** PyInstaller & Inno Setup.

## ⚙️ Estructura del Proyecto

* `app.py`: Punto de entrada y lógica de la interfaz de usuario.
* `excel_reader.py`: Módulo de extracción y limpieza de datos (Pandas).
* `excel_generator.py`: Motor de inserción de datos en plantillas Excel y formateo.
* `pdf_generator.py`: Conversión de Excel a PDF y manejo de capas de marca de agua.
* `xml_generator.py`: Generador de archivos XML para procesos aduaneros.
* `template.xlsx`: Plantilla base con el diseño corporativo.
* `template.xlm`: Plantilla base para xml.

## 🔧 Instalación y Uso

### Para Usuarios Finales
1. Descarga el instalador desde la sección de **Releases**.
2. Ejecuta `SDW_Installer_v1.1.exe`.
3. Selecciona tu archivo Excel de entrada y presiona "Generar".

### Para Desarrolladores
1. Clona el repositorio: `git clone https://github.com/Francisco-Jara-v/SDW_Shipping_Generator.git`
2. Instala dependencias: `pip install -r requirements.txt`
3. Ejecuta la aplicación: `python app.py`

## 👨‍💻 Autor
**Francisco Jara** - *Desarrollador de Software & Arquitecto de Sistemas*
