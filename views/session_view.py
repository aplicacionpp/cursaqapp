import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from excel_model import WorkbookModel
from views.dashboard import Dashboard
from views.session_view import SessionView
from views.search_view import SearchView

class TrainingApp(tk.Tk):
    def __init__(self, excel_path=None):
        super().__init__()
        self.title("Entrenamiento MVP")
        self.geometry("1200x800")

        self.model = None
        if excel_path:
            self.load_excel(excel_path)

    def load_excel(self, path):
        try:
            self.model = WorkbookModel(path)
            self.build_ui()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar Excel:\n{e}")

    def build_ui(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)

        # Dashboard
        nb.add(Dashboard(nb, self.model), text="Dashboard")

        # Sesiones: añadir una pestaña por hoja (ejemplo: primeras 3)
        for name in self.model.sheet_names[:3]:
            nb.add(SessionView(nb, self.model, name), text=name)

        # Buscador
        nb.add(SearchView(nb, self.model), text="Buscar")

if __name__ == "__main__":
    app = TrainingApp("Maider 2025_26 1er. MESO.xlsx")
    app.mainloop()
