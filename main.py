
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os, json
from typing import Optional
from excel_model import WorkbookModel
from ui_components import SummaryView, SheetPreview, FormulaSearchView, StatusBar

APP_TITLE = "Demo Excel Entrenamiento - Analizador"

class MainApp(tk.Tk):
    def __init__(self, initial_path: Optional[str] = None):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1200x800")
        self.minsize(1000, 700)

        self.model: Optional[WorkbookModel] = None
        self.current_center = None

        self._build_menu()
        self.status = StatusBar(self)
        self.status.pack(side="bottom", fill="x")

        if initial_path and os.path.exists(initial_path):
            self.load_workbook(initial_path)

    def _build_menu(self):
        m = tk.Menu(self)
        # Archivo
        filem = tk.Menu(m, tearoff=0)
        filem.add_command(label="Abrir Excel…", command=self.on_open)
        filem.add_separator()
        filem.add_command(label="Exportar resumen JSON…", command=self.on_export_json, state="disabled")
        filem.add_separator()
        filem.add_command(label="Salir", command=self.destroy)
        m.add_cascade(label="Archivo", menu=filem)

        # Ver
        viewm = tk.Menu(m, tearoff=0)
        viewm.add_command(label="Resumen de hojas", command=self.show_summary, state="disabled")
        viewm.add_command(label="Vista previa de hoja…", command=self.show_sheet_preview, state="disabled")
        viewm.add_command(label="Buscar fórmulas", command=self.show_formula_search, state="disabled")
        m.add_cascade(label="Ver", menu=viewm)

        # Ayuda
        helpm = tk.Menu(m, tearoff=0)
        helpm.add_command(label="Acerca de", command=self.on_about)
        m.add_cascade(label="Ayuda", menu=helpm)

        self.config(menu=m)
        self._menu = {"file": filem, "view": viewm}

    def enable_model_menus(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        self._menu["file"].entryconfig("Exportar resumen JSON…", state=state)
        self._menu["view"].entryconfig("Resumen de hojas", state=state)
        self._menu["view"].entryconfig("Vista previa de hoja…", state=state)
        self._menu["view"].entryconfig("Buscar fórmulas", state=state)

    def on_open(self):
        path = filedialog.askopenfilename(
            title="Selecciona un archivo .xlsx",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xltx *.xltm")]
        )
        if path:
            self.load_workbook(path)

    def load_workbook(self, path: str):
        try:
            self.status.set(f"Cargando {path} …")
            self.model = WorkbookModel(path)
            self.status.set(f"Cargado: {os.path.basename(path)} • {len(self.model.sheet_names)} hojas")
            self.enable_model_menus(True)
            self.show_summary()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el Excel:\n{e}")
            self.status.set("Error al cargar el Excel.")

    def _set_center(self, widget: tk.Widget):
        if self.current_center:
            self.current_center.destroy()
        self.current_center = widget
        widget.pack(fill="both", expand=True)

    def show_summary(self):
        if not self.model:
            return
        self._set_center(SummaryView(self, self.model))

    def show_sheet_preview(self):
        if not self.model:
            return
        # Selector simple
        win = tk.Toplevel(self)
        win.title("Elegir hoja")
        box = ttk.Combobox(win, values=self.model.sheet_names, state="readonly", width=60)
        box.pack(padx=12, pady=12)
        def open_preview(*_):
            name = box.get()
            if not name:
                return
            self._set_center(SheetPreview(self, self.model, name))
            win.destroy()
        ttk.Button(win, text="Abrir", command=open_preview).pack(padx=12, pady=(0,12))
        box.bind("<<ComboboxSelected>>", open_preview)

    def show_formula_search(self):
        if not self.model:
            return
        self._set_center(FormulaSearchView(self, self.model))

    def on_export_json(self):
        if not self.model:
            return
        path = filedialog.asksaveasfilename(
            title="Guardar resumen JSON",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")]
        )
        if path:
            out = self.model.export_summary_json(path)
            self.status.set(f"Resumen guardado en {out}")
            messagebox.showinfo("Exportado", f"Se guardó el resumen en:\n{out}")

    def on_about(self):
        messagebox.showinfo(
            "Acerca de",
            "Demo de lectura de Excel (fichero de entrenamiento):\n"
            "- Explora hojas, celdas y fórmulas\n"
            "- Muestra valores en caché (Excel) y no recalcula fórmulas\n"
            "- Permite buscar por fragmentos de fórmula\n"
            "\nConstruido con Tkinter + openpyxl."
        )

if __name__ == "__main__":
    # Si existe el Excel del usuario en el mismo directorio, lo abre automáticamente
    DEFAULT = os.environ.get("EXCEL_DEMO_DEFAULT", "Maider 2025_26 1er. MESO.xlsx")
    initial = DEFAULT if os.path.exists(DEFAULT) else None
    app = MainApp(initial_path=initial)
    app.mainloop()
