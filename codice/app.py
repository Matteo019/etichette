"""
app.py — Interfaccia grafica Tkinter per il generatore di etichette.
Eseguire con:  python app.py
"""

import os
import sys
import glob
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Importa la logica dal modulo principale
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from genera_etichette import run


# ─── Palette colori ───────────────────────────────────────────────────────────
BG        = "#1e1e2e"   # sfondo scuro
BG2       = "#2a2a3e"   # pannello secondario
ACCENT    = "#7c6af7"   # viola primario
ACCENT_H  = "#9b8df9"   # viola hover
FG        = "#cdd6f4"   # testo principale
FG_DIM    = "#6c7086"   # testo disattivato
SUCCESS   = "#a6e3a1"   # verde
WARNING   = "#f9e2af"   # giallo
ERROR     = "#f38ba8"   # rosso
FONT_MAIN = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 10, "bold")
FONT_MONO = ("Consolas", 9)
FONT_HEAD = ("Segoe UI", 16, "bold")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Generatore Etichette Aziendali")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(720, 560)

        self._build_ui()
        self._auto_detect_excel()

        # Centra la finestra
        self.update_idletasks()
        w, h = 820, 620
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ── Costruzione UI ────────────────────────────────────────────────────────

    def _build_ui(self):
        # Header
        header = tk.Frame(self, bg=ACCENT, pady=14)
        header.pack(fill="x")
        tk.Label(header, text="🏷  Generatore Etichette", font=FONT_HEAD,
                 bg=ACCENT, fg="white").pack()
        tk.Label(header, text="Formato griglia 19×4 · EAN-13 · PDF pronto per la fresa",
                 font=("Segoe UI", 9), bg=ACCENT, fg="#e0d9ff").pack()

        # Body
        body = tk.Frame(self, bg=BG, padx=24, pady=18)
        body.pack(fill="both", expand=True)

        # ── Sezione Excel ────────────────────────────────────────────────
        self._section_label(body, "📂  File Excel")

        row_excel = tk.Frame(body, bg=BG)
        row_excel.pack(fill="x", pady=(4, 0))

        self.var_excel = tk.StringVar()
        entry_excel = tk.Entry(row_excel, textvariable=self.var_excel,
                               font=FONT_MAIN, bg=BG2, fg=FG,
                               insertbackground=FG, relief="flat",
                               highlightthickness=1, highlightbackground=FG_DIM,
                               highlightcolor=ACCENT)
        entry_excel.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))

        btn_browse = self._btn(row_excel, "Sfoglia…", self._browse_excel, width=10)
        btn_browse.pack(side="left")

        # ── Sezione foglio ───────────────────────────────────────────────
        self._section_label(body, "📋  Nome foglio Excel")

        self.var_sheet = tk.StringVar(value="Situazione Magazzino")
        entry_sheet = tk.Entry(body, textvariable=self.var_sheet,
                               font=FONT_MAIN, bg=BG2, fg=FG,
                               insertbackground=FG, relief="flat",
                               highlightthickness=1, highlightbackground=FG_DIM,
                               highlightcolor=ACCENT)
        entry_sheet.pack(fill="x", ipady=6, pady=(4, 0))

        # ── Pulsante Genera ──────────────────────────────────────────────
        btn_frame = tk.Frame(body, bg=BG)
        btn_frame.pack(fill="x", pady=18)

        self.btn_genera = self._btn(btn_frame, "▶  Genera PDF", self._avvia,
                                    big=True, width=22)
        self.btn_genera.pack(side="left")

        self.btn_apri = self._btn(btn_frame, "📄  Apri PDF", self._apri_pdf,
                                  width=14, secondary=True)
        self.btn_apri.pack(side="left", padx=(12, 0))
        self.btn_apri.config(state="disabled")

        # ── Barra di progresso ───────────────────────────────────────────
        self._section_label(body, "⏳  Progresso")

        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("Custom.Horizontal.TProgressbar",
                        troughcolor=BG2, background=ACCENT,
                        borderwidth=0, thickness=10)

        self.progress_var = tk.IntVar(value=0)
        self.progress_bar = ttk.Progressbar(body, variable=self.progress_var,
                                             maximum=100,
                                             style="Custom.Horizontal.TProgressbar")
        self.progress_bar.pack(fill="x", pady=(4, 0))

        self.lbl_status = tk.Label(body, text="In attesa…", font=FONT_MAIN,
                                   bg=BG, fg=FG_DIM, anchor="w")
        self.lbl_status.pack(fill="x", pady=(4, 10))

        # ── Log ──────────────────────────────────────────────────────────
        self._section_label(body, "📝  Log elaborazione")

        log_frame = tk.Frame(body, bg=BG2, highlightthickness=1,
                             highlightbackground=FG_DIM)
        log_frame.pack(fill="both", expand=True, pady=(4, 0))

        self.txt_log = tk.Text(log_frame, font=FONT_MONO, bg=BG2, fg=FG,
                               insertbackground=FG, relief="flat",
                               state="disabled", wrap="word",
                               padx=8, pady=6)
        self.txt_log.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(log_frame, command=self.txt_log.yview,
                                 bg=BG2, troughcolor=BG2,
                                 highlightthickness=0, relief="flat")
        scrollbar.pack(side="right", fill="y")
        self.txt_log.config(yscrollcommand=scrollbar.set)

        # Tag colori log
        self.txt_log.tag_config("warn",    foreground=WARNING)
        self.txt_log.tag_config("ok",      foreground=SUCCESS)
        self.txt_log.tag_config("err",     foreground=ERROR)
        self.txt_log.tag_config("default", foreground=FG)

    # ── Helpers UI ────────────────────────────────────────────────────────────

    def _section_label(self, parent, text):
        tk.Label(parent, text=text, font=FONT_BOLD,
                 bg=BG, fg=ACCENT_H, anchor="w").pack(fill="x", pady=(14, 0))

    def _btn(self, parent, text, cmd, width=None, big=False, secondary=False):
        bg_c = BG2 if secondary else ACCENT
        fg_c = FG  if secondary else "white"
        fnt  = ("Segoe UI", 11, "bold") if big else FONT_BOLD

        b = tk.Button(parent, text=text, command=cmd,
                      bg=bg_c, fg=fg_c, activebackground=ACCENT_H,
                      activeforeground="white", relief="flat",
                      font=fnt, cursor="hand2",
                      padx=18, pady=8)
        if width:
            b.config(width=width)

        b.bind("<Enter>", lambda e: b.config(bg=ACCENT_H if not secondary else ACCENT))
        b.bind("<Leave>", lambda e: b.config(bg=bg_c))
        return b

    # ── Logica ────────────────────────────────────────────────────────────────

    def _auto_detect_excel(self):
        """Cerca automaticamente il primo .xlsx nella cartella padre."""
        parent_dir = os.path.abspath(
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
        )
        xlsx_files = glob.glob(os.path.join(parent_dir, "*.xlsx"))
        if xlsx_files:
            self.var_excel.set(xlsx_files[0])
            self._log(f"File Excel rilevato automaticamente: {xlsx_files[0]}", "ok")

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Seleziona il file Excel",
            filetypes=[("File Excel", "*.xlsx *.xls"), ("Tutti i file", "*.*")]
        )
        if path:
            self.var_excel.set(path)

    def _log(self, msg, tag="default"):
        def _do():
            self.txt_log.config(state="normal")
            # Scegli tag in base al contenuto se non specificato
            if tag == "default":
                if "ATTENZIONE" in msg or "Errore" in msg:
                    t = "warn"
                elif "successo" in msg.lower() or "✔" in msg:
                    t = "ok"
                else:
                    t = "default"
            else:
                t = tag
            self.txt_log.insert("end", msg + "\n", t)
            self.txt_log.see("end")
            self.txt_log.config(state="disabled")
        self.after(0, _do)

    def _set_status(self, msg, color=FG_DIM):
        self.after(0, lambda: self.lbl_status.config(text=msg, fg=color))

    def _set_progress(self, current, total):
        pct = int(current / total * 100) if total else 0
        self.after(0, lambda: self.progress_var.set(pct))
        self.after(0, lambda: self._set_status(f"Elaborazione: {current}/{total} righe ({pct}%)", FG))

    def _avvia(self):
        excel_path = self.var_excel.get().strip()
        sheet_name = self.var_sheet.get().strip()

        if not excel_path:
            messagebox.showwarning("File mancante", "Seleziona prima un file Excel.")
            return
        if not os.path.exists(excel_path):
            messagebox.showerror("File non trovato", f"Il file non esiste:\n{excel_path}")
            return

        # Resetta UI
        self.txt_log.config(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.config(state="disabled")
        self.progress_var.set(0)
        self.btn_genera.config(state="disabled")
        self.btn_apri.config(state="disabled")
        self._set_status("Avvio elaborazione…", FG)
        self._pdf_path = None

        # Esegui in thread separato
        t = threading.Thread(
            target=self._run_thread,
            args=(excel_path, sheet_name),
            daemon=True
        )
        t.start()

    def _run_thread(self, excel_path, sheet_name):
        try:
            result = run(
                excel_path=excel_path,
                sheet_name=sheet_name,
                log_cb=self._log,
                progress_cb=self._set_progress,
            )
            if result:
                self._pdf_path = result
                self._log(f"\n✔ PDF generato: {result}", "ok")
                self._set_status(f"Completato! → {result}", SUCCESS)
                self.after(0, lambda: self.btn_apri.config(state="normal"))
            else:
                self._log("Elaborazione terminata senza PDF.", "warn")
                self._set_status("Nessun PDF generato.", WARNING)
        except Exception as e:
            self._log(f"ERRORE CRITICO: {e}", "err")
            self._set_status(f"Errore: {e}", ERROR)
        finally:
            self.after(0, lambda: self.btn_genera.config(state="normal"))
            self.after(0, lambda: self.progress_var.set(100))

    def _apri_pdf(self):
        if self._pdf_path and os.path.exists(self._pdf_path):
            os.startfile(self._pdf_path)
        else:
            messagebox.showinfo("PDF non trovato", "Il file PDF non è disponibile.")


if __name__ == "__main__":
    app = App()
    app.mainloop()
