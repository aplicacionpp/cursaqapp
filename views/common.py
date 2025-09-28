import tkinter as tk
from tkinter import ttk
from excel_model import WorkbookModel

class Dashboard(ttk.Frame):
    def __init__(self, master, model: WorkbookModel):
        super().__init__(master, padding=10)
        self.model = model
        ttk.Label(self, text="🏋️ Dashboard Entrenamiento", font=("TkDefaultFont", 14, "bold")).pack(anchor="w")

        stats = self.model.all_summaries()
        total_cells = sum(s.nonempty_cells for s in stats)
        total_formulas = sum(s.formula_count for s in stats)

        info = [
            ("Número de hojas", len(stats)),
            ("Celdas con datos", total_cells),
            ("Total de fórmulas", total_formulas),
        ]

        tree = ttk.Treeview(self, columns=("metric", "value"), show="headings", height=len(info))
        tree.heading("metric", text="Métrica")
        tree.heading("value", text="Valor")
        for m, v in info:
            tree.insert("", "end", values=(m, v))
        tree.pack(fill="x", pady=10)
