# Generatore di Etichette Aziendali

Questo progetto permette di generare automaticamente etichette per il magazzino in formato PDF (con codici a barre EAN-13) a partire da un file Excel (.xlsx).

## Utilizzo Rapido (Standalone)

La via più semplice per usare il programma è tramite l'eseguibile portatile (per Windows):
1. Vai nella cartella `dist/` e avvia `GeneratoreEtichette.exe`.
2. Si aprirà l'interfaccia grafica.
3. Seleziona il tuo file Excel e la cartella contenente le immagini dei prodotti.
4. (Opzionale) Premi **Avvia Anteprima** per visualizzare come verrà formattata la pagina (con funzioni di zoom e pan del mouse).
5. Premi **Avvia Generazione** per creare il PDF finale. 

Tutto l'avanzamento verrà mostrato nella barra di progresso e nella console di Log visiva.

## Struttura del Progetto per Sviluppatori
- `codice/app.py`: Script che avvia l'interfaccia grafica (GUI) in stile nativo Windows tramite Tkinter.
- `codice/genera_etichette.py`: Logica core per la costruzione dei barcode e l'impaginazione in PDF.
- `GeneratoreEtichette.spec`: File di configurazione per `PyInstaller`.
- `dist/`: Contiene l'eseguibile finale pacchettizzato.

## Sviluppo e Installazione manuale

> **Prerequisito**: Python installato sul sistema.

1. Crea l'ambiente virtuale:
   ```powershell
   python -m venv .venv
   ```
2. Attiva l'ambiente virtuale:
   ```powershell
   & ".\.venv\Scripts\Activate.ps1"
   ```
3. Installa le librerie necessarie:
   ```powershell
   python -m pip install pillow python-barcode img2pdf pandas openpyxl
   ```
4. Lancia l'interfaccia grafica:
   ```bash
   python codice/app.py
   ```

## Compilazione (Pacchettizzazione in .exe)
Per generare l'eseguibile indipendente:
```powershell
pyinstaller --noconsole --onefile --name "GeneratoreEtichette" --clean --hidden-import openpyxl --hidden-import barcode.ean --hidden-import barcode.writer codice/app.py
```
