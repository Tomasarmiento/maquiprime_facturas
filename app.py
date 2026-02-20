from __future__ import annotations

import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

from processor import Processor

DEFAULT_EXCEL_NAME = "FICHERO_CONTROL_2026.xlsx"
VERSION = "v1.1"

# ── Brand colours ──────────────────────────────────────────────────────────────
BLUE        = "#1A3BFF"
BLUE_DARK   = "#1229CC"
BLUE_DEEP   = "#0A1660"
BLUE_BG     = "#EEF1FF"
WHITE       = "#FFFFFF"
OFF_WHITE   = "#F4F6FF"
GRAY_100    = "#EEF0F8"
GRAY_300    = "#C8CCDE"
GRAY_500    = "#7880A0"
GRAY_700    = "#3D4260"
TEXT        = "#111827"
SUCCESS     = "#16A34A"
WARNING_C   = "#D97706"
ERROR_C     = "#DC2626"


def _load_logo() -> tk.PhotoImage | None:
    """Load logo.png from same folder as script."""
    here = Path(__file__).parent
    for name in ("logo.png", "LOGOS-variante color-07.png"):
        p = here / name
        if p.exists():
            try:
                img = tk.PhotoImage(file=str(p))
                # subsample to ~34px height
                orig_h = img.height()
                factor = max(1, orig_h // 34)
                if factor > 1:
                    img = img.subsample(factor, factor)
                return img
            except Exception:
                pass
    return None


class HoverButton(tk.Label):
    """Label styled as a clickable button with hover effect."""

    def __init__(self, parent, text, command=None,
                 bg=BLUE, fg=WHITE, hover_bg=BLUE_DARK,
                 font=("Helvetica", 10, "bold"),
                 padx=14, pady=6, **kwargs):
        super().__init__(parent, text=text, bg=bg, fg=fg,
                         font=font, padx=padx, pady=pady,
                         cursor="hand2", **kwargs)
        self._bg = bg
        self._hover_bg = hover_bg
        self._command = command
        self._enabled = True
        self.bind("<Enter>",    self._on_enter)
        self.bind("<Leave>",    self._on_leave)
        self.bind("<Button-1>", self._on_click)

    def _on_enter(self, _=None):
        if self._enabled:
            self.config(bg=self._hover_bg)

    def _on_leave(self, _=None):
        if self._enabled:
            self.config(bg=self._bg)

    def _on_click(self, _=None):
        if self._enabled and self._command:
            self._command()

    def set_enabled(self, enabled: bool):
        self._enabled = enabled
        if enabled:
            self.config(bg=self._bg, cursor="hand2")
        else:
            self.config(bg=GRAY_300, cursor="")


class AnimatedBar(tk.Canvas):
    """Smooth indeterminate progress bar."""
    BAR_W = 140

    def __init__(self, parent, **kwargs):
        kwargs.pop("bg", None)
        super().__init__(parent, height=4, bg=GRAY_100,
                         highlightthickness=0, **kwargs)
        self._pos = 0
        self._running = False
        self._job = None

    def start(self):
        self._running = True
        self._pos = -self.BAR_W
        self._tick()

    def stop(self):
        self._running = False
        if self._job:
            self.after_cancel(self._job)
            self._job = None
        self.delete("all")

    def _tick(self):
        if not self._running:
            return
        w = self.winfo_width() or 800
        self.delete("all")
        x0, x1 = self._pos, self._pos + self.BAR_W
        self.create_rectangle(x0, 0, x1, 4, fill=BLUE, outline="")
        self._pos += 10
        if self._pos > w:
            self._pos = -self.BAR_W
        self._job = self.after(16, self._tick)


class InvoiceApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Procesador de Facturas · MAQUIPRIME")
        self.geometry("880x700")
        self.minsize(760, 600)
        self.configure(bg=OFF_WHITE)

        self.base_var    = tk.StringVar()
        self.excel_var   = tk.StringVar()
        self.dry_run_var = tk.BooleanVar(value=False)

        self._stat_inserted = tk.StringVar(value="—")
        self._stat_warnings = tk.StringVar(value="—")
        self._stat_errors   = tk.StringVar(value="—")

        self._build_ui()

    # ── UI ─────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_header()
        body = tk.Frame(self, bg=OFF_WHITE)
        body.pack(fill="both", expand=True, padx=28, pady=20)
        body.columnconfigure(0, weight=1)
        self._build_paths_card(body)
        self._build_options_card(body)
        self._build_stats_row(body)
        self._build_action_area(body)
        self._build_log_area(body)
        self._build_footer()

    # header ───────────────────────────────────────────────────────────────────

    def _build_header(self):
        hdr = tk.Frame(self, bg=WHITE, height=62)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        inner = tk.Frame(hdr, bg=WHITE)
        inner.pack(fill="both", expand=True, padx=28)

        logo_img = _load_logo()
        if logo_img:
            lbl = tk.Label(inner, image=logo_img, bg=WHITE)
            lbl.image = logo_img
            lbl.pack(side="left", pady=14)
        else:
            tk.Label(inner, text="MAQUIPRIME",
                     font=("Helvetica", 15, "bold"),
                     fg=BLUE, bg=WHITE).pack(side="left", pady=14)

        tk.Label(inner, text="PROCESADOR DE FACTURAS",
                 font=("Helvetica", 8, "bold"),
                 fg=GRAY_500, bg=WHITE).pack(side="right", pady=14)

        # blue accent line
        tk.Frame(self, bg=BLUE, height=3).pack(fill="x")

    # helpers ──────────────────────────────────────────────────────────────────

    def _card(self, parent) -> tk.Frame:
        card = tk.Frame(parent, bg=WHITE,
                        highlightthickness=1,
                        highlightbackground=GRAY_100)
        card.pack(fill="x", pady=(0, 12))
        return card

    def _section_label(self, parent, text):
        tk.Label(parent, text=text,
                 font=("Helvetica", 8, "bold"),
                 fg=GRAY_500, bg=WHITE).pack(anchor="w", padx=20, pady=(14, 8))

    def _divider(self, parent):
        tk.Frame(parent, bg=GRAY_100, height=1).pack(fill="x", padx=20)

    # paths card ───────────────────────────────────────────────────────────────

    def _build_paths_card(self, parent):
        card = self._card(parent)
        self._section_label(card, "RUTAS DE ACCESO")
        self._divider(card)

        self._field_row(card, "Carpeta base (2026)", self.base_var,
                        self._pick_base,
                        extra=("Autodetectar Excel", self._autodetect_excel))
        self._divider(card)
        self._field_row(card, "Archivo Excel central", self.excel_var,
                        self._pick_excel)

        tk.Label(card,
                 text="  ℹ  Selecciona la carpeta 2026 de tu Dropbox local sincronizado",
                 font=("Helvetica", 8), fg=GRAY_500, bg=WHITE
                 ).pack(anchor="w", padx=20, pady=(0, 12))

    def _field_row(self, card, label_text, var, browse_cmd, extra=None):
        row = tk.Frame(card, bg=WHITE)
        row.pack(fill="x", padx=20, pady=10)
        row.columnconfigure(1, weight=1)

        tk.Label(row, text=label_text,
                 font=("Helvetica", 10), fg=GRAY_700, bg=WHITE,
                 width=24, anchor="w").grid(row=0, column=0, sticky="w")

        entry = tk.Entry(row, textvariable=var,
                         font=("Helvetica", 10), fg=TEXT, bg=OFF_WHITE,
                         relief="flat",
                         highlightthickness=1,
                         highlightbackground=GRAY_300,
                         highlightcolor=BLUE,
                         readonlybackground=OFF_WHITE,
                         state="readonly")
        entry.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        btn_frame = tk.Frame(row, bg=WHITE)
        btn_frame.grid(row=0, column=2)

        HoverButton(btn_frame, text="Seleccionar", command=browse_cmd,
                    bg=WHITE, fg=BLUE, hover_bg=BLUE_BG,
                    font=("Helvetica", 9, "bold"),
                    relief="flat",
                    highlightthickness=1,
                    highlightbackground=BLUE,
                    padx=10, pady=5).pack(side="left", padx=(0, 4))

        if extra:
            HoverButton(btn_frame, text=extra[0], command=extra[1],
                        bg=WHITE, fg=GRAY_700, hover_bg=GRAY_100,
                        font=("Helvetica", 9),
                        relief="flat",
                        highlightthickness=1,
                        highlightbackground=GRAY_300,
                        padx=10, pady=5).pack(side="left")

    # options card ─────────────────────────────────────────────────────────────

    def _build_options_card(self, parent):
        card = self._card(parent)
        self._section_label(card, "OPCIONES")
        self._divider(card)

        row = tk.Frame(card, bg=WHITE)
        row.pack(fill="x", padx=20, pady=12)

        self._toggle_cv = tk.Canvas(row, width=44, height=24,
                                    bg=WHITE, highlightthickness=0)
        self._toggle_cv.pack(side="left", padx=(0, 12))
        self._draw_toggle(False)
        self._toggle_cv.bind("<Button-1>", self._on_toggle_click)

        tk.Label(row, text="Modo simulación",
                 font=("Helvetica", 10, "bold"),
                 fg=GRAY_700, bg=WHITE).pack(side="left")
        tk.Label(row, text="  —  No modifica el Excel, solo muestra qué procesaría",
                 font=("Helvetica", 9), fg=GRAY_500, bg=WHITE).pack(side="left")

        tk.Frame(card, bg=WHITE, height=2).pack()

    def _draw_toggle(self, on: bool):
        c = self._toggle_cv
        c.delete("all")
        track_color = BLUE if on else GRAY_300
        c.create_oval(0, 2, 44, 22, fill=track_color, outline="")
        knob_x = 31 if on else 13
        c.create_oval(knob_x - 9, 3, knob_x + 9, 21,
                      fill=WHITE, outline="")

    def _on_toggle_click(self, _=None):
        val = not self.dry_run_var.get()
        self.dry_run_var.set(val)
        self._draw_toggle(val)

    # stats row ────────────────────────────────────────────────────────────────

    def _build_stats_row(self, parent):
        row = tk.Frame(parent, bg=OFF_WHITE)
        row.pack(fill="x", pady=(0, 12))

        for i, (label, var, color) in enumerate([
            ("PROCESADAS",   self._stat_inserted, BLUE),
            ("ADVERTENCIAS", self._stat_warnings, WARNING_C),
            ("ERRORES",      self._stat_errors,   SUCCESS),
        ]):
            card = tk.Frame(row, bg=WHITE,
                            highlightthickness=1,
                            highlightbackground=GRAY_100)
            card.pack(side="left", fill="both", expand=True,
                      padx=(0, 8) if i < 2 else 0)

            tk.Frame(card, bg=color, height=3).pack(fill="x")
            tk.Label(card, textvariable=var,
                     font=("Helvetica", 30, "bold"),
                     fg=color, bg=WHITE).pack(pady=(10, 2))
            tk.Label(card, text=label,
                     font=("Helvetica", 8, "bold"),
                     fg=GRAY_500, bg=WHITE).pack(pady=(0, 12))

    # action area ──────────────────────────────────────────────────────────────

    def _build_action_area(self, parent):
        area = tk.Frame(parent, bg=OFF_WHITE)
        area.pack(fill="x", pady=(0, 12))

        self.process_btn = HoverButton(
            area, text="▶   PROCESAR FACTURAS",
            command=self._start_processing,
            bg=BLUE, fg=WHITE, hover_bg=BLUE_DARK,
            font=("Helvetica", 13, "bold"),
            padx=0, pady=14, relief="flat",
        )
        self.process_btn.pack(fill="x")

        self.progress_bar = AnimatedBar(area, bg=GRAY_100)
        self.progress_bar.pack(fill="x", pady=(6, 0))

    # log area ─────────────────────────────────────────────────────────────────

    def _build_log_area(self, parent):
        log_card = tk.Frame(parent, bg=BLUE_DEEP)
        log_card.pack(fill="both", expand=True)

        log_hdr = tk.Frame(log_card, bg="#0D1F80")
        log_hdr.pack(fill="x")
        tk.Label(log_hdr, text="ACTIVIDAD",
                 font=("Helvetica", 8, "bold"),
                 fg="#AAAACC", bg="#0D1F80").pack(side="left", padx=16, pady=8)
        tk.Label(log_hdr, text="●",
                 font=("Helvetica", 9),
                 fg=SUCCESS, bg="#0D1F80").pack(side="right", padx=16)

        self.log_text = tk.Text(
            log_card,
            font=("Courier", 9),
            bg=BLUE_DEEP, fg="#AAAACC",
            relief="flat",
            wrap="word",
            padx=16, pady=12,
            state="disabled",
            cursor="arrow",
        )
        self.log_text.pack(fill="both", expand=True)
        self.log_text.tag_config("ok",   foreground="#4ade80")
        self.log_text.tag_config("warn", foreground="#fbbf24")
        self.log_text.tag_config("err",  foreground="#f87171")
        self.log_text.tag_config("time", foreground="#5566AA")
        self.log_text.tag_config("info", foreground="#C8CCEE")

    # footer ───────────────────────────────────────────────────────────────────

    def _build_footer(self):
        tk.Frame(self, bg=GRAY_100, height=1).pack(fill="x", side="bottom")
        ft = tk.Frame(self, bg=WHITE, height=34)
        ft.pack(fill="x", side="bottom")
        ft.pack_propagate(False)
        inner = tk.Frame(ft, bg=WHITE)
        inner.pack(fill="both", expand=True, padx=28)
        tk.Label(inner, text="MAQUIPRIME · Gastos MX 2026",
                 font=("Helvetica", 8), fg=GRAY_500, bg=WHITE
                 ).pack(side="left", pady=8)
        tk.Label(inner, text=f"  {VERSION}  ",
                 font=("Helvetica", 8, "bold"),
                 fg=BLUE, bg=BLUE_BG).pack(side="right", pady=8)

    # ── Logging ────────────────────────────────────────────────────────────────

    def _log(self, msg: str):
        now = datetime.now().strftime("%H:%M:%S")
        msg_up = msg.upper()
        if "ERROR" in msg_up:
            tag = "err"
        elif "ADVERTENCIA" in msg_up or "WARNING" in msg_up:
            tag = "warn"
        elif any(k in msg_up for k in ("✓", "FINALIZADO", "COMPLETADO", "AUTODETECTADO")):
            tag = "ok"
        else:
            tag = "info"

        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{now}] ", "time")
        self.log_text.insert("end", f"{msg}\n", tag)
        self.log_text.configure(state="disabled")
        self.log_text.see("end")
        self.update_idletasks()

    # ── File pickers ───────────────────────────────────────────────────────────

    def _pick_base(self):
        sel = filedialog.askdirectory(title="Selecciona carpeta 2026")
        if sel:
            self.base_var.set(sel)
            self._autodetect_excel()

    def _pick_excel(self):
        sel = filedialog.askopenfilename(
            title=f"Selecciona {DEFAULT_EXCEL_NAME}",
            filetypes=[("Excel", "*.xlsx")],
        )
        if sel:
            self.excel_var.set(sel)

    def _autodetect_excel(self):
        base = self.base_var.get().strip()
        if not base:
            self._log("Primero selecciona la carpeta base 2026.")
            return
        candidate = Path(base) / DEFAULT_EXCEL_NAME
        if candidate.exists():
            self.excel_var.set(str(candidate))
            self._log(f"✓ Excel autodetectado: {candidate.name}")
        else:
            self._log(f"No se encontró {DEFAULT_EXCEL_NAME} en la carpeta seleccionada.")

    # ── Processing ─────────────────────────────────────────────────────────────

    def _start_processing(self):
        base  = Path(self.base_var.get().strip())
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

        self.process_btn.set_enabled(False)
        self.progress_bar.start()
        self._stat_inserted.set("…")
        self._stat_warnings.set("…")
        self._stat_errors.set("…")
        self._log("Iniciando procesamiento...")

        threading.Thread(
            target=self._process,
            args=(base, excel, self.dry_run_var.get()),
            daemon=True,
        ).start()

    def _process(self, base: Path, excel: Path, dry_run: bool):
        try:
            result = Processor(base, excel, self._log).run(dry_run=dry_run)

            self._stat_inserted.set(str(result["inserted"]))
            self._stat_warnings.set(str(result["warnings"]))
            self._stat_errors.set(str(result["errors"]))

            mode = " [SIMULACIÓN]" if dry_run else ""
            self._log(
                f"✓ Finalizado{mode} — "
                f"Insertadas: {result['inserted']} | "
                f"Advertencias: {result['warnings']} | "
                f"Errores: {result['errors']}"
            )
            messagebox.showinfo(
                "Completado",
                f"Procesamiento finalizado.{mode}\n\n"
                f"Insertadas:    {result['inserted']}\n"
                f"Advertencias:  {result['warnings']}\n"
                f"Errores:       {result['errors']}",
            )
        except Exception as exc:
            self._log(f"ERROR crítico: {exc}")
            messagebox.showerror("Error crítico", str(exc))
        finally:
            self.progress_bar.stop()
            self.process_btn.set_enabled(True)


if __name__ == "__main__":
    app = InvoiceApp()
    app.mainloop()