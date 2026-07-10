# Generatore di Etichette Aziendali

Questo progetto permette di generare automaticamente etichette per il magazzino in formato PDF (con codici a barre EAN-13) a partire da un file Excel (.xlsx).

## Struttura del Progetto
- `codice/`: Contiene lo script principale (`genera_etichette.py`) e il file delle dipendenze (`requirements.txt`).
- `codice/Immagini_articoli_Sigla/`: Cartella in cui inserire le immagini dei prodotti scaricate e richiamate dall'Excel.

## Installazione

> **Prerequisito**: Python non è nel PATH di sistema. Usa il percorso completo o aggiungi Python al PATH.

1. Dalla cartella principale (`etichette/`), crea l'ambiente virtuale con il percorso completo di Python:
   ```powershell
   C:\Users\_matt32\AppData\Local\Python\bin\python.exe -m venv .venv
   ```
2. Attiva l'ambiente virtuale (PowerShell su Windows):
   ```powershell
   & ".\.venv\Scripts\Activate.ps1"
   ```
3. Installa le librerie necessarie:
   ```powershell
   python -m pip install pillow python-barcode img2pdf pandas openpyxl
   ```

## Utilizzo

1. Inserisci il tuo file con i dati (`.xlsx`) nella directory principale (quella in cui si trova questo README).
2. Assicurati che le eventuali immagini dei prodotti siano all'interno di `codice/Immagini_articoli_Sigla/`.
3. Lancia lo script (ricordati di avere l'ambiente virtuale attivo):
   ```bash
   cd codice
   python genera_etichette.py
   ```
4. A processo terminato, il file `etichette_finali.pdf` pronto per la stampa verrà salvato automaticamente nella directory principale. (I file temporanei PNG generati durante il processo verranno puliti in automatico).
