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
from PIL import Image, ImageTk

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from genera_etichette import run, genera_label_anteprima

import pandas as pd

# ─── Palette colori ───────────────────────────────────────────────────────────
BG       = "#1e1e2e"
BG2      = "#2a2a3e"
BG3      = "#313244"
ACCENT   = "#7c6af7"
ACCENT_H = "#9b8df9"
FG       = "#cdd6f4"
FG_DIM   = "#6c7086"
SUCCESS  = "#a6e3a1"
WARNING  = "#f9e2af"
ERROR    = "#f38ba8"
BORDER   = "#45475a"

FONT_MAIN = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 10, "bold")
FONT_MONO = ("Consolas", 9)
FONT_HEAD = ("Segoe UI", 15, "bold")
FONT_SUB  = ("Segoe UI", 8)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Generatore Etichette Aziendali")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(980, 640)
        self._pdf_path   = None
        self._preview_tk = None   # riferimento per evitare GC

        self._build_ui()
        self._try_auto_detect_excel()

        self.update_idletasks()
        w, h = 1100, 720
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ─────────────────────────────────────────────────────────────────────────
    # BUILD UI
    # ─────────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Header ──────────────────────────────────────────────────────────
        header = tk.Frame(self, bg=ACCENT, pady=12)
        header.pack(fill="x")
        tk.Label(header, text="🏷  Generatore Etichette Aziendali",
                 font=FONT_HEAD, bg=ACCENT, fg="white").pack()
        tk.Label(header, text="Griglia 19 × 4  ·  EAN-13  ·  PDF pronto per la fresa",
                 font=("Segoe UI", 9), bg=ACCENT, fg="#ddd8ff").pack()

        # ── Main body: colonna sx + colonna dx ──────────────────────────────
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=0, pady=0)
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        # Colonna sinistra (parametri + log)
        left = tk.Frame(body, bg=BG, padx=20, pady=16)
        left.grid(row=0, column=0, sticky="nsew")

        # Colonna destra (anteprima)
        right = tk.Frame(body, bg=BG2, padx=16, pady=16,
                         highlightthickness=1, highlightbackground=BORDER)
        right.grid(row=0, column=1, sticky="nsew")

        self._build_left(left)
        self._build_right(right)

    # ─── Colonna sinistra ─────────────────────────────────────────────────────

    def _build_left(self, parent):
        # ── File Excel ──────────────────────────────────────────────────────
        self._sec(parent, "📂  File Excel *")
        row1 = tk.Frame(parent, bg=BG); row1.pack(fill="x", pady=(4,0))
        self.var_excel = tk.StringVar()
        self.var_excel.trace_add("write", self._on_params_change)
        self._entry(row1, self.var_excel).pack(side="left", fill="x", expand=True, ipady=6, padx=(0,8))
        self._btn(row1, "Sfoglia…", self._browse_excel, width=9).pack(side="left")

        # ── Cartella immagini ────────────────────────────────────────────────
        self._sec(parent, "🖼  Cartella immagini *")
        row2 = tk.Frame(parent, bg=BG); row2.pack(fill="x", pady=(4,0))
        self.var_imgdir = tk.StringVar()
        self.var_imgdir.trace_add("write", self._on_params_change)
        self._entry(row2, self.var_imgdir).pack(side="left", fill="x", expand=True, ipady=6, padx=(0,8))
        self._btn(row2, "Sfoglia…", self._browse_imgdir, width=9).pack(side="left")

        # ── Nome foglio ──────────────────────────────────────────────────────
        self._sec(parent, "📋  Nome foglio Excel")
        self.var_sheet = tk.StringVar(value="Situazione Magazzino")
        self._entry(parent, self.var_sheet).pack(fill="x", ipady=6, pady=(4,0))

        # ── Gestione anomalie ────────────────────────────────────────────────
        self._sec(parent, "⚠️  Gestione anomalie")

        anom_frame = tk.Frame(parent, bg=BG3, padx=12, pady=10,
                              highlightthickness=1, highlightbackground=BORDER)
        anom_frame.pack(fill="x", pady=(4,0))

        # Barcode
        tk.Label(anom_frame, text="Barcode mancante / invalido:",
                 font=FONT_BOLD, bg=BG3, fg=FG).pack(anchor="w")
        self.var_anom_bc = tk.StringVar(value="blank")
        self._radio_row(anom_frame, self.var_anom_bc, [
            ("blank", "Genera etichetta con zona barcode bianca"),
            ("skip",  "Salta l'etichetta"),
        ])

        tk.Frame(anom_frame, bg=BORDER, height=1).pack(fill="x", pady=8)

        # Immagine
        tk.Label(anom_frame, text="Immagine non trovata:",
                 font=FONT_BOLD, bg=BG3, fg=FG).pack(anchor="w")
        self.var_anom_img = tk.StringVar(value="blank")
        self._radio_row(anom_frame, self.var_anom_img, [
            ("blank", "Genera etichetta con metà sinistra bianca"),
            ("skip",  "Salta l'etichetta"),
        ])

        # ── Bottoni principali ───────────────────────────────────────────────
        btn_row = tk.Frame(parent, bg=BG); btn_row.pack(fill="x", pady=14)

        self.btn_avvia = self._btn(btn_row, "▶  Avvia generazione",
                                   self._avvia, big=True, width=22)
        self.btn_avvia.pack(side="left")
        self.btn_avvia.config(state="disabled")

        self.btn_apri = self._btn(btn_row, "📄  Apri PDF", self._apri_pdf,
                                  width=12, secondary=True)
        self.btn_apri.pack(side="left", padx=(10,0))
        self.btn_apri.config(state="disabled")

        # ── Progresso ────────────────────────────────────────────────────────
        self._sec(parent, "⏳  Progresso")
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("C.Horizontal.TProgressbar",
                        troughcolor=BG2, background=ACCENT,
                        borderwidth=0, thickness=8)
        self.progress_var = tk.IntVar(value=0)
        ttk.Progressbar(parent, variable=self.progress_var,
                        maximum=100, style="C.Horizontal.TProgressbar"
                        ).pack(fill="x", pady=(4,0))
        self.lbl_status = tk.Label(parent, text="Seleziona i parametri obbligatori (*) per abilitare l'avvio.",
                                   font=FONT_MAIN, bg=BG, fg=FG_DIM, anchor="w", wraplength=500)
        self.lbl_status.pack(fill="x", pady=(4,8))

        # ── Log ──────────────────────────────────────────────────────────────
        self._sec(parent, "📝  Log")
        log_wrap = tk.Frame(parent, bg=BG2, highlightthickness=1, highlightbackground=BORDER)
        log_wrap.pack(fill="both", expand=True, pady=(4,0))
        self.txt_log = tk.Text(log_wrap, font=FONT_MONO, bg=BG2, fg=FG,
                               insertbackground=FG, relief="flat",
                               state="disabled", wrap="word", padx=8, pady=6)
        self.txt_log.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(log_wrap, command=self.txt_log.yview,
                          bg=BG2, troughcolor=BG2, highlightthickness=0, relief="flat")
        sb.pack(side="right", fill="y")
        self.txt_log.config(yscrollcommand=sb.set)
        self.txt_log.tag_config("warn", foreground=WARNING)
        self.txt_log.tag_config("ok",   foreground=SUCCESS)
        self.txt_log.tag_config("err",  foreground=ERROR)

    # ─── Colonna destra (anteprima) ───────────────────────────────────────────

    def _build_right(self, parent):
        tk.Label(parent, text="🔍  Anteprima etichetta",
                 font=FONT_BOLD, bg=BG2, fg=ACCENT_H).pack(anchor="w")
        tk.Label(parent, text="Generata dalla prima riga valida del file Excel",
                 font=FONT_SUB, bg=BG2, fg=FG_DIM).pack(anchor="w", pady=(0,10))

        self.canvas_preview = tk.Canvas(parent, bg=BG3, highlightthickness=1,
                                        highlightbackground=BORDER)
        self.canvas_preview.pack(fill="both", expand=True)
        self.lbl_preview_info = tk.Label(parent,
                                         text="Seleziona il file Excel per generare l'anteprima.",
                                         font=FONT_MAIN, bg=BG2, fg=FG_DIM,
                                         anchor="center", wraplength=300)
        self.lbl_preview_info.pack(pady=8)

        self.btn_refresh = self._btn(parent, "🔄  Aggiorna anteprima",
                                     self._refresh_preview, secondary=True)
        self.btn_refresh.pack(pady=(0,4))
        self.btn_refresh.config(state="disabled")

    # ─────────────────────────────────────────────────────────────────────────
    # HELPERS UI
    # ─────────────────────────────────────────────────────────────────────────

    def _sec(self, parent, text):
        tk.Label(parent, text=text, font=FONT_BOLD,
                 bg=BG, fg=ACCENT_H, anchor="w").pack(fill="x", pady=(14,0))

    def _entry(self, parent, var):
        return tk.Entry(parent, textvariable=var, font=FONT_MAIN,
                        bg=BG2, fg=FG, insertbackground=FG, relief="flat",
                        highlightthickness=1, highlightbackground=BORDER,
                        highlightcolor=ACCENT)

    def _btn(self, parent, text, cmd, width=None, big=False, secondary=False):
        bg_c = BG3 if secondary else ACCENT
        fg_c = FG  if secondary else "white"
        fnt  = ("Segoe UI", 11, "bold") if big else FONT_BOLD
        b = tk.Button(parent, text=text, command=cmd,
                      bg=bg_c, fg=fg_c, activebackground=ACCENT_H,
                      activeforeground="white", relief="flat",
                      font=fnt, cursor="hand2", padx=14, pady=7)
        if width:
            b.config(width=width)
        b.bind("<Enter>", lambda e: b.config(bg=ACCENT_H if not secondary else ACCENT))
        b.bind("<Leave>", lambda e: b.config(bg=bg_c))
        return b

    def _radio_row(self, parent, var, options):
        for value, label in options:
            tk.Radiobutton(parent, text=label, variable=var, value=value,
                           font=FONT_MAIN, bg=BG3, fg=FG,
                           selectcolor=ACCENT, activebackground=BG3,
                           activeforeground=FG, cursor="hand2"
                           ).pack(anchor="w", padx=8, pady=1)

    # ─────────────────────────────────────────────────────────────────────────
    # LOGICA
    # ─────────────────────────────────────────────────────────────────────────

    def _try_auto_detect_excel(self):
        parent_dir = os.path.abspath(
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
        )
        xlsx_files = glob.glob(os.path.join(parent_dir, "*.xlsx"))
        if xlsx_files:
            self.var_excel.set(xlsx_files[0])
        # Cerca anche cartella immagini default
        default_img = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   "Immagini_articoli_Sigla")
        if os.path.isdir(default_img):
            self.var_imgdir.set(default_img)

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Seleziona il file Excel",
            filetypes=[("File Excel", "*.xlsx *.xls"), ("Tutti i file", "*.*")]
        )
        if path:
            self.var_excel.set(path)

    def _browse_imgdir(self):
        path = filedialog.askdirectory(title="Seleziona la cartella delle immagini")
        if path:
            self.var_imgdir.set(path)

    def _on_params_change(self, *_):
        """Abilita/disabilita il tasto Avvia in base ai parametri obbligatori."""
        excel_ok = bool(self.var_excel.get().strip()) and os.path.isfile(self.var_excel.get().strip())
        img_ok   = bool(self.var_imgdir.get().strip()) and os.path.isdir(self.var_imgdir.get().strip())

        state = "normal" if (excel_ok and img_ok) else "disabled"
        self.btn_avvia.config(state=state)
        self.btn_refresh.config(state=state)

        if excel_ok and img_ok:
            self._set_status("Parametri pronti. Premi ▶ Avvia per generare il PDF.", FG)
            self._refresh_preview()
        else:
            missing = []
            if not excel_ok:
                missing.append("file Excel")
            if not img_ok:
                missing.append("cartella immagini")
            self._set_status(f"Seleziona: {', '.join(missing)} per abilitare l'avvio.", FG_DIM)

    def _refresh_preview(self):
        """Genera l'anteprima della prima riga valida in un thread."""
        excel  = self.var_excel.get().strip()
        imgdir = self.var_imgdir.get().strip()
        sheet  = self.var_sheet.get().strip()
        if not (excel and imgdir):
            return
        threading.Thread(target=self._preview_thread,
                         args=(excel, imgdir, sheet), daemon=True).start()

    def _preview_thread(self, excel, imgdir, sheet):
        try:
            df = pd.read_excel(excel, sheet_name=sheet)
            pil_img = None
            for idx, row in df.iterrows():
                pil_img = genera_label_anteprima(row, imgdir)
                if pil_img:
                    break
            if pil_img:
                self.after(0, lambda img=pil_img: self._show_preview(img))
            else:
                self.after(0, lambda: self.lbl_preview_info.config(
                    text="Nessuna riga valida trovata per l'anteprima."))
        except Exception as e:
            self.after(0, lambda: self.lbl_preview_info.config(
                text=f"Errore anteprima: {e}"))

    def _show_preview(self, pil_img):
        self.canvas_preview.update_idletasks()
        cw = self.canvas_preview.winfo_width()
        ch = self.canvas_preview.winfo_height()
        if cw < 10 or ch < 10:
            cw, ch = 400, 200

        ratio = min(cw / pil_img.width, ch / pil_img.height)
        nw    = max(1, int(pil_img.width  * ratio))
        nh    = max(1, int(pil_img.height * ratio))
        resized = pil_img.resize((nw, nh), Image.Resampling.LANCZOS)

        self._preview_tk = ImageTk.PhotoImage(resized)
        self.canvas_preview.delete("all")
        self.canvas_preview.create_image(cw // 2, ch // 2, anchor="center",
                                         image=self._preview_tk)
        self.lbl_preview_info.config(text=f"Anteprima: {pil_img.width}×{pil_img.height} px")

    # ── Generazione PDF ───────────────────────────────────────────────────────

    def _avvia(self):
        excel  = self.var_excel.get().strip()
        imgdir = self.var_imgdir.get().strip()
        sheet  = self.var_sheet.get().strip()

        if not excel or not os.path.isfile(excel):
            messagebox.showerror("File non trovato", f"Il file Excel non esiste:\n{excel}")
            return
        if not imgdir or not os.path.isdir(imgdir):
            messagebox.showerror("Cartella non trovata",
                                 f"La cartella immagini non esiste:\n{imgdir}")
            return

        # Reset UI
        self.txt_log.config(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.config(state="disabled")
        self.progress_var.set(0)
        self.btn_avvia.config(state="disabled")
        self.btn_apri.config(state="disabled")
        self._set_status("Avvio elaborazione…", FG)
        self._pdf_path = None

        threading.Thread(
            target=self._run_thread,
            args=(excel, imgdir, sheet,
                  self.var_anom_bc.get(), self.var_anom_img.get()),
            daemon=True
        ).start()

    def _run_thread(self, excel, imgdir, sheet, anom_bc, anom_img):
        try:
            result = run(
                excel_path=excel,
                sheet_name=sheet,
                immagini_dir=imgdir,
                anomalia_barcode=anom_bc,
                anomalia_immagine=anom_img,
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
            self.after(0, lambda: self.btn_avvia.config(state="normal"))
            self.after(0, lambda: self.progress_var.set(100))

    def _apri_pdf(self):
        if self._pdf_path and os.path.exists(self._pdf_path):
            os.startfile(self._pdf_path)
        else:
            messagebox.showinfo("PDF non trovato", "Il file PDF non è disponibile.")

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _log(self, msg, tag=None):
        def _do():
            t = tag
            if t is None:
                if "ATTENZIONE" in msg or "Errore" in msg or "ERRORE" in msg:
                    t2 = "warn"
                elif "✔" in msg or "successo" in msg.lower():
                    t2 = "ok"
                else:
                    t2 = None
            else:
                t2 = t
            self.txt_log.config(state="normal")
            self.txt_log.insert("end", msg + "\n", t2 or "")
            self.txt_log.see("end")
            self.txt_log.config(state="disabled")
        self.after(0, _do)

    def _set_status(self, msg, color=FG_DIM):
        self.after(0, lambda: self.lbl_status.config(text=msg, fg=color))

    def _set_progress(self, current, total):
        pct = int(current / total * 100) if total else 0
        self.after(0, lambda: self.progress_var.set(pct))
        self.after(0, lambda: self._set_status(
            f"Elaborazione: {current}/{total} righe ({pct}%)", FG))


if __name__ == "__main__":
    app = App()
    app.mainloop()
