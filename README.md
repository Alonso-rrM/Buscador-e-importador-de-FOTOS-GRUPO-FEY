# Unificador de PDF y Exportador de Imágenes (Catálogo)

Una aplicación de escritorio moderna desarrollada en **Python** con **Tkinter** para automatizar el procesamiento de catálogos y la generación de documentos PDF unificados, optimizada para entornos corporativos con accesos a servidores de archivos compartidos (`SMB/LAN`).

Este proyecto fue diseñado con un enfoque profesional, implementando una interfaz de usuario atractiva en **modo oscuro** (Material Dark Theme), manejo de base de datos master en formato CSV (`pandas`) y empaquetado final a ejecutable (`.exe`) independiente para facilitar su distribución.

## 🚀 Características Principales

*   **Generador de PDF Personalizado:** Busca y unifica imágenes de productos desde un servidor de archivos en base a códigos ingresados, organizándolas en un PDF formateado.
*   **Exportador de Imágenes:** Copia de forma rápida imágenes individuales a una carpeta seleccionada por el usuario, evitando tener que buscarlas manualmente en directorios de red complejos.
*   **Filtros Inteligentes:** Reglas automáticas para procesar códigos de artículos específicos (por ejemplo, artículos con prefijo `xz` y descarte automático de variantes no deseadas).
*   **Base de Datos en Red Integrada:** Conexión y lectura optimizada en segundos de un archivo maestro de base de datos (`.db` / CSV de gran volumen con más de 200,000 registros) ubicado en un servidor local de la empresa (`\\feyserver`).
*   **Interfaz de Usuario Material Dark:** Una interfaz limpia, moderna y responsiva con soporte para:
    *   Limpia completa de entradas en un solo clic.
    *   Ruta predeterminada editable para la base de datos local o de red.
    *   Manejo de errores detallado (por ejemplo, fallos de credenciales de red `WinError 1326`).
*   **Suite de Pruebas Automatizadas:** Incluye pruebas unitarias integrales (`verify.py`) para validar el correcto funcionamiento de los filtros, lógica de búsqueda, ordenamiento por fecha de modificación, y el generador de PDF.

## 🛠️ Tecnologías y Librerías Utilizadas

*   **Lenguaje:** Python 3.11+
*   **Interfaz Gráfica:** Tkinter (con estilos modernos y manejo de hilos para no congelar la GUI durante tareas pesadas)
*   **Procesamiento de Datos:** [Pandas](https://pandas.pydata.org/) (carga rápida e indexación O(1) de la base de datos)
*   **Manipulación de Imágenes:** [Pillow (PIL)](https://python-pillow.org/)
*   **Automatización de Compilación:** [PyInstaller](https://pyinstaller.org/) (empaquetado a un único archivo ejecutable `.exe` sin consola)

## 📁 Estructura del Proyecto

*   `unificador.py`: Módulo que contiene la lógica central del negocio: lectura e indexación del CSV, generación del PDF y exportación de imágenes.
*   `gui.py`: Interfaz gráfica de usuario con diseño oscuro, validación de campos y llamadas asíncronas para evitar bloqueos de la ventana.
*   `verify.py`: Conjunto de pruebas unitarias robustas que validan la lógica de filtrado y generación de archivos.
*   `generate_test_data.py`: Script para crear datos de prueba artificiales que permiten probar la aplicación de forma aislada.
*   `build.bat`: Script automatizado para compilar la aplicación a un archivo `.exe` en la carpeta `dist/`.
*   `.gitignore`: Excluye de manera segura entornos virtuales, archivos temporales de compilación y datos de pruebas locales.

## 💻 Cómo Ejecutar el Proyecto en Desarrollo

1.  **Clona el repositorio:**
    ```bash
    git clone <URL_DE_TU_REPOSITORIO>
    cd pdf_unifier
    ```

2.  **Crea y activa un entorno virtual:**
    ```bash
    python -m venv .venv
    # En Windows:
    .venv\Scripts\activate
    ```

3.  **Instala las dependencias:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Ejecuta la interfaz gráfica:**
    ```bash
    python gui.py
    ```

5.  **Ejecuta las pruebas unitarias:**
    ```bash
    python verify.py
    ```

## 📦 Compilación a Ejecutable (.exe)

Para generar la versión ejecutable distribuible para Windows, puedes ejecutar el archivo por lotes en la consola:
```bash
build.bat
```
Esto creará el archivo ejecutable optimizado en `dist/UnificadorPDF.exe`.

---
*Desarrollado como parte de soluciones de automatización de TI y portafolio profesional.*
