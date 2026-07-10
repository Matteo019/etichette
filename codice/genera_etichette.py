"""
genera_etichette.py — Logica di generazione etichette EAN-13 in formato PDF.
Può essere usato da CLI (python genera_etichette.py) o importato dall'UI (app.py).
"""

import os
import sys
import glob
import shutil
import argparse
import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageOps
import barcode
from barcode.writer import ImageWriter
import img2pdf


# ─────────────────────────────────────────────────────────────────────────────
# Costanti layout etichetta (750×373 px)
# ─────────────────────────────────────────────────────────────────────────────
W, H         = 750, 373
CENTER_X     = int(W * 0.731)
MAX_TXT_W    = int(W * 0.385)
LEFT_BOX_W   = int(W * 0.48)

SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
FONT_BOLD    = os.path.join(SCRIPT_DIR, "arialbd.ttf")
FONT_REGULAR = os.path.join(SCRIPT_DIR, "arial.ttf")


# ─────────────────────────────────────────────────────────────────────────────
# Helper: estrai stringa sicura da cella Excel
# ─────────────────────────────────────────────────────────────────────────────

def _safe(val):
    if pd.isna(val) or str(val).strip().lower() == "nan":
        return ""
    return str(val).strip()


# ─────────────────────────────────────────────────────────────────────────────
# Genera una singola etichetta PIL Image da una riga del DataFrame
# ─────────────────────────────────────────────────────────────────────────────

def genera_label(row, immagini_dir,
                 anomalia_barcode="blank", anomalia_immagine="blank",
                 temp_dir=None):
    """
    Genera un'etichetta PIL Image da una riga del DataFrame.

    Args:
        row:               riga pandas (Series o dict-like)
        immagini_dir:      percorso cartella immagini locali
        anomalia_barcode:  "blank" → lascia bianco | "skip" → ritorna None
        anomalia_immagine: "blank" → lascia bianco | "skip" → ritorna None
        temp_dir:          cartella temporanea per i barcode (default: script_dir/temp_barcodes)

    Returns:
        PIL.Image oppure None se la riga va saltata
    """
    if temp_dir is None:
        temp_dir = os.path.join(SCRIPT_DIR, "temp_barcodes")
        os.makedirs(temp_dir, exist_ok=True)

    # ── Valori dalla riga ────────────────────────────────────────────────────
    barcode_val = _safe(row.get("BARCODE", ""))
    if barcode_val.endswith(".0"):
        barcode_val = barcode_val[:-2]

    sigla      = _safe(row.get("CODICE SIGLA", ""))
    modello    = _safe(row.get("MODELLO", "")).replace(" ", "   ")
    materiale  = _safe(row.get("MATERIALE", ""))
    dimensioni = _safe(row.get("DIMENSIONI", ""))
    immagine   = _safe(row.get("IMMAGINE", ""))

    # ── Immagine prodotto ────────────────────────────────────────────────────
    path_img = None
    if immagine:
        nome_img = immagine.replace("\\", "/").split("/")[-1]
        candidate = os.path.join(immagini_dir, nome_img)
        if os.path.exists(candidate):
            path_img = candidate
        else:
            if anomalia_immagine == "skip":
                return None

    # ── Barcode ──────────────────────────────────────────────────────────────
    bc_invalido = len(barcode_val) != 13 or not barcode_val.isdigit()
    if bc_invalido and anomalia_barcode == "skip":
        return None

    # ── Costruzione immagine ──────────────────────────────────────────────────
    img  = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    # Immagine prodotto (metà sinistra)
    if path_img:
        try:
            prod = Image.open(path_img).convert("RGBA")
            bg   = Image.new("RGBA", prod.size, "white")
            bg.paste(prod, (0, 0), prod)
            bg = bg.convert("RGB")
            tw, th = LEFT_BOX_W - 20, H - 20
            final = ImageOps.pad(bg, (tw, th), color="white")
            px = (LEFT_BOX_W - tw) // 2
            py = (H - th) // 2
            img.paste(final, (px, py))
        except Exception:
            pass  # lato sx resta bianco

    # Barcode
    if not bc_invalido:
        try:
            EAN = barcode.get_barcode_class("ean13")
            opts = {
                "dpi": 300, "module_width": 0.25, "module_height": 6.0,
                "font_size": 7, "text_distance": 4.5,
                "guard_height_factor": 1.2, "quiet_zone": 2.0,
                "font_path": FONT_REGULAR,
            }
            tmp_path = os.path.join(temp_dir, f"bc_{os.getpid()}")
            ean = EAN(barcode_val, writer=ImageWriter(), guardbar=True)
            testo = f"{ean.ean[0]} {ean.ean[1:7]} {ean.ean[7:]}"
            ean.save(tmp_path, options=opts, text=testo)
            bc_img = Image.open(f"{tmp_path}.png")
            bc_w, _ = bc_img.size
            img.paste(bc_img, (CENTER_X - bc_w // 2, int(H * 0.056)))
            bc_img.close()
            os.remove(f"{tmp_path}.png")
        except Exception:
            pass  # zona barcode resta bianca

    # Testi centrati
    def draw_text(text, h_pct, y_pct):
        if not text:
            return
        try:
            fnt  = ImageFont.truetype(FONT_BOLD, int(H * h_pct))
            bbox = fnt.getbbox(text)
            tw   = bbox[2] - bbox[0]
            th   = bbox[3] - bbox[1]
            tmp  = Image.new("RGBA", (tw + 10, th + 20), (255, 255, 255, 0))
            td   = ImageDraw.Draw(tmp)
            td.text((5, 10 - bbox[1]), text, fill="black", font=fnt)
            rb = tmp.getbbox()
            if rb:
                tmp = tmp.crop(rb)
            nw = tmp.width
            if nw > MAX_TXT_W:
                nw = MAX_TXT_W
                tmp = tmp.resize((nw, tmp.height), Image.Resampling.LANCZOS)
            mask = tmp.split()[3]
            img.paste(tmp, (CENTER_X - nw // 2, int(H * y_pct)), mask)
        except Exception:
            pass

    draw_text(sigla,      0.10, 0.390)
    draw_text(modello,    0.25, 0.505)
    draw_text(materiale,  0.12, 0.755)
    draw_text(dimensioni, 0.11, 0.880)

    return img


def genera_label_anteprima(row, immagini_dir):
    """Wrapper per l'anteprima GUI: usa sempre 'blank' per non saltare nulla."""
    return genera_label(row, immagini_dir,
                        anomalia_barcode="blank",
                        anomalia_immagine="blank")


# ─────────────────────────────────────────────────────────────────────────────
# Funzione principale: genera tutte le etichette e assembla il PDF
# ─────────────────────────────────────────────────────────────────────────────

def run(excel_path,
        sheet_name="Situazione Magazzino",
        immagini_dir=None,
        anomalia_barcode="blank",
        anomalia_immagine="blank",
        log_cb=None,
        progress_cb=None):
    """
    Genera il PDF delle etichette.

    Args:
        excel_path:        percorso al file .xlsx
        sheet_name:        nome del foglio
        immagini_dir:      cartella delle immagini (default: ./Immagini_articoli_Sigla)
        anomalia_barcode:  "blank" | "skip"
        anomalia_immagine: "blank" | "skip"
        log_cb:            callback(str, tag=None) per log UI
        progress_cb:       callback(current, total) per progress bar

    Returns:
        str: percorso PDF, oppure None
    """
    def log(msg, tag=None):
        if log_cb:
            log_cb(msg, tag)
        else:
            print(msg)

    if immagini_dir is None:
        immagini_dir = os.path.join(SCRIPT_DIR, "Immagini_articoli_Sigla")

    output_dir = os.path.join(SCRIPT_DIR, "output_etichette")
    temp_dir   = os.path.join(SCRIPT_DIR, "temp_barcodes")
    pdf_path   = os.path.abspath(os.path.join(SCRIPT_DIR, "..", "etichette_finali.pdf"))

    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(temp_dir,   exist_ok=True)

    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
    except Exception as e:
        log(f"Errore lettura Excel: {e}", "err")
        return None

    total = len(df)
    log(f"Inizio elaborazione di {total} righe...")

    generated_images = []

    for idx, row in df.iterrows():
        excel_row   = idx + 2
        barcode_val = _safe(row.get("BARCODE", ""))
        if barcode_val.endswith(".0"):
            barcode_val = barcode_val[:-2]

        # Log anomalie prima della generazione
        bc_invalido = len(barcode_val) != 13 or not barcode_val.isdigit()
        if bc_invalido:
            log(f"ATTENZIONE (Riga {excel_row}): Barcode invalido ({barcode_val})"
                f" → {'saltata' if anomalia_barcode == 'skip' else 'zona bianca'}", "warn")

        immagine = _safe(row.get("IMMAGINE", ""))
        if immagine:
            nome_img  = immagine.replace("\\", "/").split("/")[-1]
            candidate = os.path.join(immagini_dir, nome_img)
            if not os.path.exists(candidate):
                log(f"ATTENZIONE (Riga {excel_row}): Immagine {nome_img} non trovata"
                    f" → {'saltata' if anomalia_immagine == 'skip' else 'lato bianco'}", "warn")

        # Genera etichetta
        pil_img = genera_label(row, immagini_dir,
                               anomalia_barcode=anomalia_barcode,
                               anomalia_immagine=anomalia_immagine,
                               temp_dir=temp_dir)

        if pil_img is None:
            # Riga saltata per policy anomalie
            if progress_cb:
                progress_cb(idx + 1, total)
            continue

        filename = f"etichetta_{idx+1}.png"
        out_path = os.path.join(output_dir, filename)
        pil_img.save(out_path)
        generated_images.append(out_path)

        if progress_cb:
            progress_cb(idx + 1, total)

    log(f"\nGenerate {len(generated_images)} etichette su {total} righe.")

    # ── Creazione PDF ─────────────────────────────────────────────────────────
    if not generated_images:
        log("Nessuna etichetta generata, PDF non creato.", "warn")
        shutil.rmtree(temp_dir,   ignore_errors=True)
        shutil.rmtree(output_dir, ignore_errors=True)
        return None

    log("Creazione delle pagine a griglia in corso...")
    cols = 19
    rows = 4
    gap  = 0
    labels_per_page = cols * rows

    w_rot = H
    h_rot = W

    margin_x = 70
    margin_y = 30
    grid_w   = cols * w_rot
    grid_h   = rows * h_rot
    page_w   = grid_w + 2 * margin_x
    page_h   = grid_h + 2 * margin_y

    pages = []
    for i in range(0, len(generated_images), labels_per_page):
        page_img  = Image.new("RGB", (page_w, page_h), "white")
        draw_page = ImageDraw.Draw(page_img)

        batch = generated_images[i:i+labels_per_page]
        for j, img_path in enumerate(batch):
            try:
                lbl     = Image.open(img_path)
                lbl_rot = lbl.transpose(Image.Transpose.ROTATE_90)
                col     = j % cols
                row_idx = j // cols
                x = margin_x + col     * w_rot
                y = margin_y + row_idx * h_rot
                page_img.paste(lbl_rot, (x, y))
                lbl_rot.close()
                lbl.close()
            except Exception as e:
                log(f"Errore incollaggio {img_path}: {e}", "warn")

        # Cerchi di riferimento
        mark_d   = 40
        cx_left  = margin_x / 2
        cx_right = page_w - margin_x / 2
        for mx, my in [
            (cx_left,  page_h * 0.2017),
            (cx_left,  page_h * 0.9717),
            (cx_right, page_h * 0.0901),
            (cx_right, page_h * 0.6361),
        ]:
            draw_page.ellipse(
                [mx - mark_d/2, my - mark_d/2, mx + mark_d/2, my + mark_d/2],
                outline="black", width=3
            )

        # Griglia minima
        for c in range(cols + 1):
            lx = margin_x + c * w_rot
            draw_page.line([(lx, margin_y), (lx, margin_y + grid_h)],
                           fill="black", width=1)
        for r in range(rows + 1):
            ly = margin_y + r * h_rot
            draw_page.line([(margin_x, ly), (margin_x + grid_w, ly)],
                           fill="black", width=1)

        # Squadretta a "L" in basso a destra
        bk = 3
        vx = page_w * 0.99764
        vy = page_h * 0.99356
        hl = page_w * 0.005805
        vl = page_h * 0.009442
        draw_page.line([(vx - hl, vy), (vx, vy)], fill="black", width=bk)
        draw_page.line([(vx, vy - vl), (vx, vy)], fill="black", width=bk)

        page_path = os.path.join(output_dir, f"page_grid_{i//labels_per_page + 1}.png")
        page_img.save(page_path)
        pages.append(page_path)

    log("Creazione del PDF in corso...")
    with open(pdf_path, "wb") as f:
        f.write(img2pdf.convert(pages))
    log(f"PDF salvato con successo in: {pdf_path}", "ok")

    shutil.rmtree(temp_dir,   ignore_errors=True)
    shutil.rmtree(output_dir, ignore_errors=True)

    return pdf_path


# ─────────────────────────────────────────────────────────────────────────────
# Entrypoint CLI
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Genera etichette da un file Excel.")
    parser.add_argument("excel_path", nargs="?", help="Percorso del file Excel.")
    parser.add_argument("--sheet_name",  default="Situazione Magazzino")
    parser.add_argument("--immagini_dir", default=None)
    parser.add_argument("--anomalia_barcode",  default="blank", choices=["blank","skip"])
    parser.add_argument("--anomalia_immagine", default="blank", choices=["blank","skip"])
    args = parser.parse_args()

    if args.excel_path:
        excel_path = args.excel_path
    else:
        parent_dir = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
        xlsx_files = glob.glob(os.path.join(parent_dir, "*.xlsx"))
        if not xlsx_files:
            print("Nessun file Excel trovato nella cartella superiore.")
            sys.exit(1)
        excel_path = xlsx_files[0]
        print(f"File Excel trovato: {excel_path}")

    run(excel_path,
        sheet_name=args.sheet_name,
        immagini_dir=args.immagini_dir,
        anomalia_barcode=args.anomalia_barcode,
        anomalia_immagine=args.anomalia_immagine)


if __name__ == "__main__":
    main()
