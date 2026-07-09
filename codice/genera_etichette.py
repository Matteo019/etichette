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

def create_attention_icon(width, height):
    icon_img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(icon_img)
    
    center_x = width // 2
    center_y = height // 2
    size = min(width, height) * 0.7
    
    p1 = (center_x, center_y - size // 2)
    p2 = (center_x - size // 2.2, center_y + size // 2.2)
    p3 = (center_x + size // 2.2, center_y + size // 2.2)
    
    draw.polygon([p1, p2, p3], fill="orange", outline="black", width=5)
    
    line_w = size * 0.1
    top_y = center_y - size * 0.1
    bottom_y = center_y + size * 0.15
    
    draw.line([(center_x, top_y), (center_x, bottom_y)], fill="black", width=int(line_w))
    dot_y = center_y + size * 0.3
    draw.ellipse([(center_x - line_w/2, dot_y - line_w/2), (center_x + line_w/2, dot_y + line_w/2)], fill="black")
    
    return icon_img

def main():
    parser = argparse.ArgumentParser(description="Genera etichette da un file Excel.")
    parser.add_argument("excel_path", nargs="?", help="Percorso del file Excel. Se non specificato, cerca il primo .xlsx nella directory padre.")
    parser.add_argument("--sheet_name", default="Situazione Magazzino", help="Nome del foglio Excel da leggere.")
    args = parser.parse_args()

    if args.excel_path:
        excel_path = args.excel_path
    else:
        parent_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
        xlsx_files = glob.glob(os.path.join(parent_dir, "*.xlsx"))
        if not xlsx_files:
            print("Nessun file Excel trovato nella cartella superiore.")
            sys.exit(1)
        excel_path = xlsx_files[0]
        print(f"File Excel trovato: {excel_path}")

    sheet_name = args.sheet_name
    
    output_dir = "output_etichette"
    immagini_locali_dir = "Immagini_articoli_Sigla"
    
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    
    temp_dir = "temp_barcodes"
    os.makedirs(temp_dir, exist_ok=True)
    
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"Errore lettura Excel: {e}")
        sys.exit(1)
        
    w, h = 750, 373
    font_path = "arialbd.ttf"
        
    center_x = int(w * 0.731)
    max_text_width = int(w * 0.385)
    
    generated_images = []
    
    print(f"Inizio elaborazione di {len(df)} righe...")
    
    for idx, row in df.iterrows():
        excel_row = idx + 2
        barcode_val = str(row.get('BARCODE', '')).strip()
        if barcode_val.endswith('.0'):
            barcode_val = barcode_val[:-2]
            
        sigla = str(row.get('CODICE SIGLA', '')).strip()
        modello = str(row.get('MODELLO', '')).strip().replace(' ', '   ')
        materiale = str(row.get('MATERIALE', '')).strip()
        dimensioni = str(row.get('DIMENSIONI', '')).strip()
        immagine_rete = str(row.get('IMMAGINE', '')).strip()
        
        path_img_locale = None
        if immagine_rete and immagine_rete.lower() != 'nan':
            nome_img = immagine_rete.replace("\\", "/").split("/")[-1]
            temp_path = os.path.join(immagini_locali_dir, nome_img)
            if os.path.exists(temp_path):
                path_img_locale = temp_path
            else:
                print(f"ATTENZIONE (Riga {excel_row}): Immagine {nome_img} non trovata.")
        
        if len(barcode_val) != 13 or not barcode_val.isdigit():
            print(f"ATTENZIONE (Riga {excel_row}): Barcode invalido ({barcode_val})")
            continue
            
        img = Image.new("RGB", (w, h), "white")
        draw = ImageDraw.Draw(img)
        
        left_box_w = int(w * 0.48)
        left_box_h = h
        
        icon = create_attention_icon(left_box_w, left_box_h)
        
        if path_img_locale:
            try:
                prod_img = Image.open(path_img_locale).convert("RGBA")
                bg_prod = Image.new("RGBA", prod_img.size, "white")
                bg_prod.paste(prod_img, (0, 0), prod_img)
                bg_prod = bg_prod.convert("RGB")
                
                target_w = left_box_w - 20
                target_h = left_box_h - 20
                final_prod_img = ImageOps.pad(bg_prod, (target_w, target_h), color="white")
                
                paste_x = (left_box_w - target_w) // 2
                paste_y = (left_box_h - target_h) // 2
                
                img.paste(final_prod_img, (paste_x, paste_y))
            except Exception as e:
                print(f"Errore caricamento immagine {path_img_locale}: {e}")
                img.paste(icon, (0, 0))
        else:
            img.paste(icon, (0, 0))

        draw.rectangle([0, 0, w-1, h-1], outline="black", width=2)
        
        EAN = barcode.get_barcode_class('ean13')
        try:
            writer_options = {
                'dpi': 300,
                'module_width': 0.25,
                'module_height': 6.0,
                'font_size': 7,
                'text_distance': 4.5,
                'guard_height_factor': 1.2,
                'quiet_zone': 2.0,
                'font_path': 'arial.ttf'
            }
            temp_bc = os.path.join(temp_dir, f"temp_bc_{excel_row}")
            ean = EAN(barcode_val, writer=ImageWriter(), guardbar=True)
            testo_barcode = f"{ean.ean[0]} {ean.ean[1:7]} {ean.ean[7:]}"
            ean.save(temp_bc, options=writer_options, text=testo_barcode)
            
            bc_img = Image.open(f"{temp_bc}.png")
            bc_w, bc_h = bc_img.size
            img.paste(bc_img, (center_x - bc_w // 2, int(h * 0.056)))
            bc_img.close()
            os.remove(f"{temp_bc}.png")
            
        except Exception as e:
            print(f"Riga {excel_row}: Errore barcode -> {e}")
            continue

        def draw_squashed_text_center(text, target_h_pct, y_pct):
            font_size = int(h * target_h_pct)
            font = ImageFont.truetype(font_path, font_size)
            bbox = font.getbbox(text)
            tw = bbox[2] - bbox[0]
            th = bbox[3] - bbox[1]
            
            temp_img = Image.new("RGBA", (tw + 10, th + 20), (255, 255, 255, 0))
            temp_draw = ImageDraw.Draw(temp_img)
            temp_draw.text((5, 10 - bbox[1]), text, fill="black", font=font)
            
            real_bbox = temp_img.getbbox()
            if real_bbox:
                temp_img = temp_img.crop(real_bbox)
                
            new_w = temp_img.width
            if new_w > max_text_width:
                new_w = max_text_width
                temp_img = temp_img.resize((new_w, temp_img.height), Image.Resampling.LANCZOS)
                
            paste_x = center_x - new_w // 2
            paste_y = int(h * y_pct)
            
            mask = temp_img.split()[3]
            img.paste(temp_img, (paste_x, paste_y), mask)
            
        draw_squashed_text_center(sigla, 0.10, 0.390)
        draw_squashed_text_center(modello, 0.25, 0.505)
        draw_squashed_text_center(materiale, 0.12, 0.755)
        draw_squashed_text_center(dimensioni, 0.11, 0.880)
        
        filename = f"etichetta_{idx+1}.png"
        out_path = os.path.join(output_dir, filename)
        img.save(out_path)
        generated_images.append(out_path)
        
    print(f"\nGenerate {len(generated_images)} etichette in '{output_dir}'.")
    
    # Creazione PDF finale
    pdf_path = "../etichette_finali.pdf"
    if generated_images:
        print("Creazione delle pagine a griglia in corso...")
        cols = 19
        rows = 4
        gap = 10
        labels_per_page = cols * rows
        
        # Etichette "in piedi" (ruotate di 90 gradi)
        w_rot = h
        h_rot = w
        
        grid_w = cols * w_rot + (cols - 1) * gap
        grid_h = rows * h_rot + (rows - 1) * gap
        
        # Margini esterni
        margin_x = 70
        margin_y = 30
        
        page_w = grid_w + 2 * margin_x
        page_h = grid_h + 2 * margin_y
        
        pages = []
        for i in range(0, len(generated_images), labels_per_page):
            # Sfondo bianco per la pagina
            page_img = Image.new("RGB", (page_w, page_h), "white")
            draw_page = ImageDraw.Draw(page_img)
            
            batch = generated_images[i:i+labels_per_page]
            
            for j, img_path in enumerate(batch):
                try:
                    lbl = Image.open(img_path)
                    # Ruotiamo l'etichetta per farla "in piedi"
                    lbl_rot = lbl.transpose(Image.Transpose.ROTATE_90)
                    col = j % cols
                    row = j // cols
                    x = margin_x + col * (w_rot + gap)
                    y = margin_y + row * (h_rot + gap)
                    page_img.paste(lbl_rot, (x, y))
                    lbl_rot.close()
                    lbl.close()
                except Exception as e:
                    print(f"Errore incollaggio {img_path}: {e}")
            
            # Disegno la linea nera esterna (tutto attorno alla pagina)
            draw_page.rectangle([0, 0, page_w-1, page_h-1], outline="black", width=2)
            
            page_path = os.path.join(output_dir, f"page_grid_{i//labels_per_page + 1}.png")
            page_img.save(page_path)
            pages.append(page_path)
            
        print("Creazione del PDF in corso...")
        with open(pdf_path, "wb") as f:
            f.write(img2pdf.convert(pages))
        print(f"PDF salvato con successo in: {pdf_path}")
    else:
        print("Nessuna etichetta generata, PDF non creato.")
        
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir, ignore_errors=True)
        
    # Pulisce le immagini png generate, lasciando solo il pdf
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir, ignore_errors=True)

if __name__ == "__main__":
    main()
