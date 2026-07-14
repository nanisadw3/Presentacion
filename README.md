# Visor de Refinación y Exportador a PowerPoint

Esta es una aplicación de escritorio desarrollada en Python con **CustomTkinter** diseñada para leer, limpiar, visualizar y exportar datos operativos de refinación (Crudo, Gasolinas, Diésel, Turbosina, Asfalto y Combustóleo, incluyendo especificaciones de la refinería Cadereyta).

## 🚀 Características Principales

*   **Interfaz Gráfica Moderna:** Construida con `CustomTkinter` y soporte para tema oscuro.
*   **Análisis y Limpieza de Excel:** Usa `pandas` para parsear automáticamente formatos rígidos de hojas de Excel, limpiando filas vacías, eliminando decimales y emparejando meses/años dinámicamente.
*   **Base de Datos Local:** Utiliza `sqlite3` (`db_helper.py`) para gestionar e inyectar registros de años extra en los cálculos anuales.
*   **Simulación Anual:** Calcula producciones, tiempos y métricas combinadas (ej. Producción x Días) para cada tipo de combustible.
*   **Exportación a PowerPoint:** Integra `python-pptx` para generar automáticamente diapositivas ejecutivas con tablas incrustadas basadas en un formato predefinido.

## 📁 Estructura del Proyecto

*   `app.py`: Controlador principal de la aplicación y la Interfaz Gráfica (GUI). Gestiona el *MainLoop*, las interacciones del usuario y orquesta la comunicación entre los otros módulos.
*   `excel_parser.py`: Motor de procesamiento de datos. Se encarga de abrir el Excel, recortar por rangos de celdas duros (`iloc`), limpiar inconsistencias y ejecutar los bucles de simulación.
*   `pptx_exporter.py`: Motor de exportación. Recibe los DataFrames procesados, carga una plantilla `.pptx` (si existe) y llena las diapositivas correspondientes.
*   `db_helper.py`: Gestión de la base de datos `historico_produccion.db` para almacenar y recuperar datos manuales de años extra.

## 🛠 Instalación y Uso

1.  Asegúrate de tener Python 3.10+ instalado.
2.  Instala las dependencias necesarias:
    ```bash
    pip install -r requirements.txt
    ```
    *(Dependencias clave: customtkinter, pandas, python-pptx, CTkTable, openpyxl)*
3.  Ejecuta la aplicación:
    ```bash
    python app.py
    ```
4.  Carga un archivo Excel válido haciendo clic en **"Cargar Excel"** y utiliza el menú desplegable para navegar entre los distintos productos.

## 🤝 Mantenimiento
Para escalar y entender el funcionamiento técnico interno de cómo leer nuevos productos, por favor revisa el archivo `.ai_context.md`.
