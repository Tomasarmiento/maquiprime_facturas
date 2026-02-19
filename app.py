from __future__ import annotations

import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from processor import Processor

DEFAULT_EXCEL_NAME = "FICHERO_CONTROL_2026.xlsx"


class InvoiceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Procesador de Facturas MAQUIPRIME")
        self.geometry("900x620")

        self.base_var = tk.StringVar()
        self.excel_var = tk.StringVar()
        self.dry_run_var = tk.BooleanVar(value=False)

        self._build_ui()

    def _build_ui(self):
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Carpeta base (2026):").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.base_var, width=95).grid(row=1, column=0, sticky="ew", padx=(0, 8))

        base_buttons = ttk.Frame(frame)
        base_buttons.grid(row=1, column=1, sticky="ns")
        ttk.Button(base_buttons, text="Seleccionar", command=self._pick_base).grid(row=0, column=0, pady=(0, 4))
        ttk.Button(base_buttons, text="Autodetectar Excel", command=self._autodetect_excel).grid(row=1, column=0)

        ttk.Label(frame, text="Archivo Excel central:").grid(row=2, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(frame, textvariable=self.excel_var, width=95).grid(row=3, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(frame, text="Seleccionar", command=self._pick_excel).grid(row=3, column=1)

        help_text = (
            "Nota: esta app NO usa URL. Debes seleccionar la ruta local sincronizada de Dropbox.\n"
            "Ejemplo: .../Joaquin GL Dropbox/MOLGROUP - URUGUAY/MAQUIPRIME/Gastos MX/2026"
        )
        ttk.Label(frame, text=help_text, foreground="#555").grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 0))

        ttk.Checkbutton(frame, text="Modo simulación (no guarda cambios)", variable=self.dry_run_var).grid(
            row=5, column=0, sticky="w", pady=(12, 0)
        )

        self.process_btn = ttk.Button(frame, text="▶ Procesar facturas", command=self._start_processing)
        self.process_btn.grid(row=6, column=0, sticky="w", pady=(12, 0))

        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(12, 0))

        ttk.Label(frame, text="Actividad:").grid(row=8, column=0, sticky="w", pady=(12, 0))
        self.log_text = tk.Text(frame, height=18, wrap="word")
        self.log_text.grid(row=9, column=0, columnspan=2, sticky="nsew")

        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(9, weight=1)

    def _pick_base(self):
        selected = filedialog.askdirectory(title="Selecciona carpeta 2026")
        if selected:
            self.base_var.set(selected)
            self._autodetect_excel()

    def _pick_excel(self):
        selected = filedialog.askopenfilename(
            title=f"Selecciona {DEFAULT_EXCEL_NAME}",
            filetypes=[("Excel", "*.xlsx")],
        )
        if selected:
            self.excel_var.set(selected)

    def _autodetect_excel(self):
        base_value = self.base_var.get().strip()
        if not base_value:
            self._log("Primero selecciona la carpeta base 2026 para autodetectar el Excel.")
            return

        candidate = Path(base_value) / DEFAULT_EXCEL_NAME
        if candidate.exists():
            self.excel_var.set(str(candidate))
            self._log(f"Excel autodetectado: {candidate}")
        else:
            self._log(f"No se encontró {DEFAULT_EXCEL_NAME} dentro de la carpeta seleccionada.")

    def _log(self, msg: str):
        now = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{now}] {msg}\n")
        self.log_text.see("end")
        self.update_idletasks()

    def _start_processing(self):
        base = Path(self.base_var.get().strip())
        excel = Path(self.excel_var.get().strip())

        if not base.exists() or not base.is_dir():
            messagebox.showerror("Error", "Selecciona una carpeta base válida.")
            return

        if excel.name != DEFAULT_EXCEL_NAME:
            messagebox.showerror("Error", f"Selecciona específicamente {DEFAULT_EXCEL_NAME}.")
            return

        if not excel.exists() or excel.suffix.lower() != ".xlsx":
            messagebox.showerror("Error", "Selecciona un archivo Excel válido (.xlsx).")
            return

        self.process_btn.configure(state="disabled")
        self.progress.start(10)
        self._log("Iniciando procesamiento...")

        t = threading.Thread(target=self._process, args=(base, excel, self.dry_run_var.get()), daemon=True)
        t.start()

    def _process(self, base: Path, excel: Path, dry_run: bool):
        try:
            processor = Processor(base, excel, self._log)
            result = processor.run(dry_run=dry_run)
            self._log(
                f"Finalizado. Insertadas: {result['inserted']} | Advertencias: {result['warnings']} | Errores: {result['errors']}"
            )
            messagebox.showinfo(
                "Completado",
                f"Procesamiento finalizado.\n\n"
                f"Insertadas: {result['inserted']}\n"
                f"Advertencias: {result['warnings']}\n"
                f"Errores: {result['errors']}",
            )
        except Exception as exc:
            self._log(f"ERROR crítico: {exc}")
            messagebox.showerror("Error crítico", str(exc))
        finally:
            self.progress.stop()
            self.process_btn.configure(state="normal")


if __name__ == "__main__":
    app = InvoiceApp()
    app.mainloop()
