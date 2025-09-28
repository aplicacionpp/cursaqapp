# 📊 cursaqapp

Aplicación de escritorio en **Python (Tkinter + openpyxl)** que toma como base un archivo Excel de planificación de entrenamientos (`Maider 2025_26 1er. MESO.xlsx`) y lo transforma en una interfaz visual:

- Permite **explorar todas las hojas** del Excel.
- Muestra **resúmenes de celdas y fórmulas**.
- Vista previa de los datos (valores cacheados).
- **Buscador de fórmulas** por texto.
- Exportación de resúmenes a **JSON** y hojas a **CSV**.

---

## 🚀 Requisitos

- **Python 3.9+** (probado en macOS y Linux, también funciona en Windows).
- Paquetes (se instalan desde `requirements.txt`):
  - `openpyxl` (manejo de Excel).
  - `Pillow` (soporte de imágenes embebidas, opcional).

Instalación de dependencias:

```bash
pip install -r requirements.txt
