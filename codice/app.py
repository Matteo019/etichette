"""
app.py — Interfaccia grafica Tkinter (Stile Classico) per il generatore di etichette.
Eseguire con:  python app.py
"""

import os
import sys
import glob
import time
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk

if getattr(sys, 'frozen', False):
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, SCRIPT_DIR)
from genera_etichette import run, genera_label_anteprima, genera_pagina_preview

import pandas as pd

class StdoutRedirector:
    def __init__(self, text_widget_queue, tag=None):
        self.queue = text_widget_queue
        self.tag = tag
    def write(self, string):
        if string.strip():
            self.queue.put((string.strip(), self.tag))
    def flush(self):
        pass

# Rate-limiting per la UI
PROGRESS_INTERVAL_MS = 150
LOG_FLUSH_INTERVAL_MS = 200

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Generatore Etichette Aziendali")
        self.geometry("1100x720")
        self.minsize(980, 640)

        # Tema nativo Windows se disponibile
        self.style = ttk.Style(self)
        try:
            if 'vista' in self.style.theme_names():
                self.style.theme_use('vista')
            elif 'xpnative' in self.style.theme_names():
                self.style.theme_use('xpnative')
        except:
            pass

        self._pdf_path        = None
        self._preview_tk      = None
        self._preview_pil     = None
        self._scale           = 1.0
        self._offset_x        = 0
        self._offset_y        = 0
        self._pan_start_x     = 0
        self._pan_start_y     = 0
        
        self._last_progress_t = 0.0
        self._log_queue       = queue.Queue()
        self._log_flush_id    = None

        sys.stdout = StdoutRedirector(self._log_queue)
        sys.stderr = StdoutRedirector(self._log_queue, tag="err")

        self._build_ui()
        self._try_auto_detect()
        self._start_log_flusher()

        self.update_idletasks()
        x = (self.winfo_screenwidth()  - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    # ─────────────────────────────────────────────────────────────────────────
    # BUILD UI
    # ─────────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Header in alto
        header = tk.Frame(self, bg="#ffffff", pady=10, padx=15, highlightthickness=1, highlightbackground="#d0d0d0")
        header.pack(fill="x")
        tk.Label(header, text="Generatore Etichette Aziendali", font=("Segoe UI", 16, "bold"), bg="#ffffff", fg="#003366", anchor="w").pack(fill="x")
        tk.Label(header, text="Griglia 19 × 4  ·  EAN-13  ·  PDF pronto per la fresa", font=("Segoe UI", 9), bg="#ffffff", fg="#555555", anchor="w").pack(fill="x")

        # Container principale diviso in 2 colonne (sinistra per controlli, destra per anteprima)
        body = tk.Frame(self, padx=10, pady=10)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        left = tk.Frame(body)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        right = tk.Frame(body)
        right.grid(row=0, column=1, sticky="nsew")

        self._build_left(left)
        self._build_right(right)

    def _build_left(self, parent):
        # ── Parametri Progetto ───────────────────────────────────────────────
        lf_params = ttk.LabelFrame(parent, text=" 📂 Parametri Progetto ")
        lf_params.pack(fill="x", pady=(0, 10), ipady=5)

        tk.Label(lf_params, text="File Excel *:").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))
        self.var_excel = tk.StringVar()
        self.var_excel.trace_add("write", self._on_params_change)
        ttk.Entry(lf_params, textvariable=self.var_excel).grid(row=0, column=1, sticky="ew", padx=5, pady=(10, 5))
        ttk.Button(lf_params, text="Sfoglia...", command=self._browse_excel, width=12).grid(row=0, column=2, padx=10, pady=(10, 5))

        tk.Label(lf_params, text="Cartella Immagini *:").grid(row=1, column=0, sticky="w", padx=10, pady=(0, 5))
        self.var_imgdir = tk.StringVar()
        self.var_imgdir.trace_add("write", self._on_params_change)
        ttk.Entry(lf_params, textvariable=self.var_imgdir).grid(row=1, column=1, sticky="ew", padx=5, pady=(0, 5))
        ttk.Button(lf_params, text="Sfoglia...", command=self._browse_imgdir, width=12).grid(row=1, column=2, padx=10, pady=(0, 5))

        tk.Label(lf_params, text="Salvataggio PDF (opzionale):").grid(row=2, column=0, sticky="w", padx=10, pady=(0, 10))
        self.var_pdf_out = tk.StringVar()
        ttk.Entry(lf_params, textvariable=self.var_pdf_out).grid(row=2, column=1, sticky="ew", padx=5, pady=(0, 10))
        ttk.Button(lf_params, text="Sfoglia...", command=self._browse_pdf, width=12).grid(row=2, column=2, padx=10, pady=(0, 10))
        
        lf_params.columnconfigure(1, weight=1)

        # ── Gestione anomalie ────────────────────────────────────────────────
        lf_anom = ttk.LabelFrame(parent, text=" ⚠️ Gestione Anomalie ")
        lf_anom.pack(fill="x", pady=10, ipady=5)

        f_bc = tk.Frame(lf_anom)
        f_bc.pack(fill="x", padx=10, pady=5)
        tk.Label(f_bc, text="Barcode mancante/invalido:").pack(side="left", padx=(0, 10))
        self.var_anom_bc = tk.StringVar(value="blank")
        ttk.Radiobutton(f_bc, text="Lascia bianco", variable=self.var_anom_bc, value="blank").pack(side="left", padx=5)
        ttk.Radiobutton(f_bc, text="Salta l'etichetta", variable=self.var_anom_bc, value="skip").pack(side="left", padx=5)

        f_img = tk.Frame(lf_anom)
        f_img.pack(fill="x", padx=10, pady=5)
        tk.Label(f_img, text="Immagine non trovata:").pack(side="left", padx=(0, 26))
        self.var_anom_img = tk.StringVar(value="blank")
        ttk.Radiobutton(f_img, text="Lascia bianco", variable=self.var_anom_img, value="blank").pack(side="left", padx=5)
        ttk.Radiobutton(f_img, text="Salta l'etichetta", variable=self.var_anom_img, value="skip").pack(side="left", padx=5)

        # ── Impostazioni avanzate ─────────────────────────────────────────────
        lf_adv = ttk.LabelFrame(parent, text=" ⚙ Impostazioni Avanzate ")
        lf_adv.pack(fill="x", pady=10, ipady=5)
        
        tk.Label(lf_adv, text="Nome foglio Excel:").pack(side="left", padx=10, pady=5)
        self.var_sheet = tk.StringVar(value="Situazione Magazzino")
        ttk.Entry(lf_adv, textvariable=self.var_sheet, width=25).pack(side="left", padx=5, pady=5)
        tk.Label(lf_adv, text="(Default: Situazione Magazzino)", fg="#666").pack(side="left", padx=5)

        # ── Azioni (Bottoni principali) ──────────────────────────────────────
        f_actions = tk.Frame(parent)
        f_actions.pack(fill="x", pady=15)
        
        self.btn_anteprima = tk.Button(f_actions, text="🔍 Avvia Anteprima", command=self._avvia_anteprima, font=("Segoe UI", 10), bg="#e1e1e1", padx=10, pady=5, relief="groove")
        self.btn_anteprima.pack(side="left")
        self.btn_anteprima.config(state="disabled")

        self.btn_avvia = tk.Button(f_actions, text="▶ Avvia Generazione", command=self._avvia, font=("Segoe UI", 10, "bold"), bg="#0066cc", fg="white", padx=10, pady=5, relief="groove")
        self.btn_avvia.pack(side="left", padx=(10, 0))
        self.btn_avvia.config(state="disabled")

        self.btn_apri = tk.Button(f_actions, text="📄 Apri PDF", command=self._apri_pdf, font=("Segoe UI", 10), bg="#e1e1e1", padx=10, pady=5, relief="groove")
        self.btn_apri.pack(side="left", padx=(10, 0))
        self.btn_apri.config(state="disabled")

        # ── Progresso ────────────────────────────────────────────────────────
        lf_prog = ttk.LabelFrame(parent, text=" ⏳ Stato ")
        lf_prog.pack(fill="x", pady=(5, 10), ipady=5)
        
        self.lbl_status = tk.Label(lf_prog, text="Seleziona i parametri obbligatori (*) per abilitare l'avvio.", fg="#666")
        self.lbl_status.pack(anchor="w", padx=10, pady=(5, 5))
        
        self.progress_var = tk.IntVar(value=0)
        ttk.Progressbar(lf_prog, variable=self.progress_var, maximum=100).pack(fill="x", padx=10, pady=(0, 5))

        # ── Log ──────────────────────────────────────────────────────────────
        lf_log = ttk.LabelFrame(parent, text=" 📝 Log Eventi ")
        lf_log.pack(fill="both", expand=True)

        self.txt_log = tk.Text(lf_log, font=("Consolas", 9), state="disabled", wrap="word", bg="white", height=8)
        self.txt_log.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        sb = ttk.Scrollbar(lf_log, command=self.txt_log.yview)
        sb.pack(side="right", fill="y", pady=5, padx=(0, 5))
        self.txt_log.config(yscrollcommand=sb.set)
        
        self.txt_log.tag_config("warn", foreground="#d97706") # arancione
        self.txt_log.tag_config("ok",   foreground="#16a34a") # verde
        self.txt_log.tag_config("err",  foreground="#dc2626") # rosso

    def _build_right(self, parent):
        lf_preview = ttk.LabelFrame(parent, text=" 🔍 Visualizzatore Anteprima ")
        lf_preview.pack(fill="both", expand=True)

        tk.Label(lf_preview, text="Anteprima visiva della pagina 4x19", fg="#666").pack(pady=(10, 5))

        self.canvas_preview = tk.Canvas(lf_preview, bg="black", highlightthickness=1, highlightbackground="#cccccc")
        self.canvas_preview.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        self.canvas_preview.bind("<MouseWheel>", self._on_zoom)
        self.canvas_preview.bind("<ButtonPress-1>", self._on_pan_start)
        self.canvas_preview.bind("<B1-Motion>", self._on_pan_move)
        
        # Gestione zoom mouse per Linux
        self.canvas_preview.bind("<Button-4>", self._on_zoom)
        self.canvas_preview.bind("<Button-5>", self._on_zoom)

        self.lbl_preview_info = tk.Label(lf_preview, text="Usa la rotellina per zoomare e trascina per spostarti.", fg="#666")
        self.lbl_preview_info.pack(pady=(0, 10))


    # ─────────────────────────────────────────────────────────────────────────
    # LOG con buffer + flush periodico
    # ─────────────────────────────────────────────────────────────────────────
    def _start_log_flusher(self):
        self._flush_log_buffer()

    def _flush_log_buffer(self):
        messages = []
        try:
            while True:
                messages.append(self._log_queue.get_nowait())
        except queue.Empty:
            pass

        if messages:
            self.txt_log.config(state="normal")
            for msg, tag in messages:
                self.txt_log.insert("end", msg + "\n", tag or "")
            self.txt_log.see("end")
            self.txt_log.config(state="disabled")

        self._log_flush_id = self.after(LOG_FLUSH_INTERVAL_MS, self._flush_log_buffer)

    def _log(self, msg, tag=None):
        if tag is None:
            if any(k in msg for k in ("ATTENZIONE", "Errore", "ERRORE")): tag = "warn"
            elif "✔" in msg or "successo" in msg.lower(): tag = "ok"
        self._log_queue.put((msg, tag))

    # ─────────────────────────────────────────────────────────────────────────
    # STATUS & PROGRESS (throttled)
    # ─────────────────────────────────────────────────────────────────────────
    def _set_status(self, msg, color="#666666"):
        self.after(0, lambda: self.lbl_status.config(text=msg, fg=color))

    def _set_progress(self, current, total):
        now = time.monotonic()
        is_last = (current >= total)
        if not is_last and (now - self._last_progress_t) < PROGRESS_INTERVAL_MS / 1000:
            return
        self._last_progress_t = now
        pct = int(current / total * 100) if total else 0

        def _update():
            self.progress_var.set(pct)
            self.lbl_status.config(text=f"Elaborazione: {current}/{total} righe ({pct}%)", fg="#000")
        self.after(0, _update)

    # ─────────────────────────────────────────────────────────────────────────
    # LOGICA APP
    # ─────────────────────────────────────────────────────────────────────────
    def _try_auto_detect(self):
        parent_dir = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
        xlsx_files = glob.glob(os.path.join(parent_dir, "*.xlsx"))
        if xlsx_files:
            self.var_excel.set(xlsx_files[0])

        default_img = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Immagini_articoli_Sigla")
        if os.path.isdir(default_img):
            self.var_imgdir.set(default_img)

    def _browse_excel(self):
        path = filedialog.askopenfilename(title="Seleziona il file Excel", filetypes=[("File Excel", "*.xlsx *.xls"), ("Tutti i file", "*.*")])
        if path:
            self.var_excel.set(path)

    def _browse_imgdir(self):
        path = filedialog.askdirectory(title="Seleziona la cartella delle immagini")
        if path:
            self.var_imgdir.set(path)

    def _browse_pdf(self):
        path = filedialog.asksaveasfilename(title="Salva PDF come", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf"), ("Tutti i file", "*.*")])
        if path:
            self.var_pdf_out.set(path)

    def _on_params_change(self, *_):
        excel_ok = bool(self.var_excel.get().strip()) and os.path.isfile(self.var_excel.get().strip())
        img_ok   = bool(self.var_imgdir.get().strip()) and os.path.isdir(self.var_imgdir.get().strip())
        ready    = excel_ok and img_ok

        state = "normal" if ready else "disabled"
        self.btn_avvia.config(state=state)
        self.btn_anteprima.config(state=state)

        if ready:
            self._set_status("Parametri pronti. Premi 🔍 Anteprima o ▶ Avvia.", "#000")
        else:
            missing = []
            if not excel_ok:  missing.append("file Excel")
            if not img_ok:    missing.append("cartella immagini")
            self._set_status(f"Seleziona: {', '.join(missing)} per abilitare l'avvio.", "#666")

    def _avvia_anteprima(self):
        excel  = self.var_excel.get().strip()
        imgdir = self.var_imgdir.get().strip()
        sheet  = self.var_sheet.get().strip()
        if not (excel and imgdir):
            return
        
        self.btn_anteprima.config(state="disabled")
        self.lbl_preview_info.config(text="Generazione anteprima in corso…", fg="#666")
        
        threading.Thread(
            target=self._pagina_preview_thread,
            args=(excel, imgdir, sheet, self.var_anom_bc.get(), self.var_anom_img.get()),
            daemon=True
        ).start()

    def _pagina_preview_thread(self, excel, imgdir, sheet, anom_bc, anom_img):
        try:
            df      = pd.read_excel(excel, sheet_name=sheet)
            pil_img = genera_pagina_preview(
                df, imgdir,
                anomalia_barcode=anom_bc,
                anomalia_immagine=anom_img,
                max_labels=76
            )
            self.after(0, lambda img=pil_img: self._show_preview(img))
            self.after(0, lambda: self.lbl_preview_info.config(text=f"Pagina 4×19 ({pil_img.width}×{pil_img.height} px)", fg="#666"))
        except Exception as e:
            self.after(0, lambda: self.lbl_preview_info.config(text=f"Errore anteprima: {e}", fg="#dc2626"))
        finally:
            self.after(0, lambda: self.btn_anteprima.config(state="normal"))

    def _show_preview(self, pil_img):
        self._preview_pil = pil_img
        self._scale = 1.0
        self._offset_x = 0
        self._offset_y = 0
        self._draw_preview()

    def _draw_preview(self):
        if not self._preview_pil:
            return
            
        self.canvas_preview.update_idletasks()
        cw = max(self.canvas_preview.winfo_width(),  10)
        ch = max(self.canvas_preview.winfo_height(), 10)
        
        # Se lo zoom è 1.0, fitta perfettamente nell'area
        base_ratio = min(cw / self._preview_pil.width, ch / self._preview_pil.height)
        ratio = base_ratio * self._scale
        
        nw, nh = max(1, int(self._preview_pil.width * ratio)), max(1, int(self._preview_pil.height * ratio))
        resized = self._preview_pil.resize((nw, nh), Image.Resampling.LANCZOS)
        self._preview_tk = ImageTk.PhotoImage(resized)
        self.canvas_preview.delete("all")
        
        center_x = (cw // 2) + self._offset_x
        center_y = (ch // 2) + self._offset_y
        self.canvas_preview.create_image(center_x, center_y, anchor="center", image=self._preview_tk)

    def _on_zoom(self, event):
        if not self._preview_pil: return
        # Windows/Mac support (event.delta), Linux (event.num)
        if event.num == 4 or getattr(event, 'delta', 0) > 0:
            self._scale *= 1.2
        elif event.num == 5 or getattr(event, 'delta', 0) < 0:
            self._scale /= 1.2
        self._draw_preview()

    def _on_pan_start(self, event):
        self.canvas_preview.focus_set()
        self._pan_start_x = event.x
        self._pan_start_y = event.y

    def _on_pan_move(self, event):
        if not self._preview_pil: return
        dx = event.x - self._pan_start_x
        dy = event.y - self._pan_start_y
        self._offset_x += dx
        self._offset_y += dy
        self._pan_start_x = event.x
        self._pan_start_y = event.y
        self._draw_preview()

    def _avvia(self):
        excel  = self.var_excel.get().strip()
        imgdir = self.var_imgdir.get().strip()
        sheet  = self.var_sheet.get().strip()

        if not excel or not os.path.isfile(excel):
            messagebox.showerror("File non trovato", f"Il file Excel non esiste:\n{excel}")
            return
        if not imgdir or not os.path.isdir(imgdir):
            messagebox.showerror("Cartella non trovata", f"La cartella immagini non esiste:\n{imgdir}")
            return

        self.txt_log.config(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.config(state="disabled")
        self.progress_var.set(0)
        self._last_progress_t = 0.0
        
        out_pdf = self.var_pdf_out.get().strip() or None
        
        self.btn_avvia.config(state="disabled")
        self.btn_apri.config(state="disabled")
        self._set_status("Avvio elaborazione…", "#000")
        self._pdf_path = None

        threading.Thread(
            target=self._run_thread,
            args=(excel, imgdir, sheet, self.var_anom_bc.get(), self.var_anom_img.get(), out_pdf),
            daemon=True
        ).start()

    def _run_thread(self, excel, imgdir, sheet, anom_bc, anom_img, out_pdf):
        try:
            result = run(
                excel_path=excel,
                sheet_name=sheet,
                immagini_dir=imgdir,
                anomalia_barcode=anom_bc,
                anomalia_immagine=anom_img,
                output_pdf_path=out_pdf,
                log_cb=self._log,
                progress_cb=self._set_progress,
            )
            if result:
                self._pdf_path = result
                self._log(f"\n✔ PDF generato: {result}", "ok")
                self._set_status(f"Completato con successo!", "#16a34a")
                self.after(0, lambda: self.btn_apri.config(state="normal"))
            else:
                self._log("Elaborazione terminata senza PDF.", "warn")
                self._set_status("Nessun PDF generato.", "#d97706")
        except Exception as e:
            self._log(f"ERRORE CRITICO: {e}", "err")
            self._set_status(f"Errore: {e}", "#dc2626")
        finally:
            self.after(0, lambda: self.btn_avvia.config(state="normal"))
            self.after(0, lambda: self.progress_var.set(100))
            self.after(1500, lambda: self.progress_var.set(0))

    def _apri_pdf(self):
        if self._pdf_path and os.path.exists(self._pdf_path):
            os.startfile(self._pdf_path)
        else:
            messagebox.showinfo("PDF non trovato", "Il file PDF non è disponibile.")


if __name__ == "__main__":
    app = App()
    app.mainloop()
