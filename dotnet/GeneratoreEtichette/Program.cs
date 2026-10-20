using System.Globalization;
using System.Text;
using SkiaSharp;

namespace GeneratoreEtichette;

static class Program
{
    const string Uso = """
        Uso:
          GeneratoreEtichette.exe <file.xlsx> <cartella immagini> [opzioni]

        Opzioni:
          --output <file.pdf>   PDF da creare (default: accanto all'Excel, "<nome> - impaginato.pdf")
          --foglio <nome>       foglio Excel da leggere (default: il primo con le colonne giuste)
          --sfondo <colore>     giallo (default), bianco oppure #RRGGBB

        Prima di creare il PDF controlla tutte le righe: se trova anche un solo errore
        (immagine mancante, barcode errato, ...) li elenca tutti e il PDF NON viene creato.
        """;

    const int LatoMinimoImmagine = 100;
    const int QuantitaSospetta = 500;

    static int Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);   // richiesto da ExcelDataReader
        try
        {
            return (int)Esegui(args);
        }
        catch (ErroreGeneratore ex)
        {
            Console.WriteLine();
            Console.WriteLine($"ERRORE: {ex.Message}");
            if (ex.Suggerimento != null)
                Console.WriteLine($"  -> {ex.Suggerimento}");
            if (ex.Esito == Esito.Parametri)
                Console.WriteLine($"{Environment.NewLine}{Uso}");
            return (int)ex.Esito;
        }
        catch (Exception ex)
        {
            Console.WriteLine();
            Console.WriteLine($"ERRORE IMPREVISTO ({ex.GetType().Name}): {ex.Message}");
            Console.WriteLine("  -> inviare questo messaggio a chi mantiene il programma. Dettagli tecnici:");
            Console.WriteLine(ex);
            return (int)Esito.Imprevisto;
        }
    }

    static Esito Esegui(string[] args)
    {
        var (excel, cartella, output, foglio, sfondo) = LeggiParametri(args);
        if (excel == null)
        {
            Console.WriteLine(Uso);
            return Esito.Ok;
        }

        Console.WriteLine($"Excel:    {excel}");
        Console.WriteLine($"Immagini: {cartella}");
        Console.WriteLine($"PDF:      {output}");

        Excel.ControllaFile(excel);
        ControllaOutput(output);
        var indice = new Immagini(cartella);
        var tabella = Excel.Leggi(excel, foglio);
        Console.WriteLine($"Foglio \"{tabella.Foglio}\": {tabella.Righe.Count} righe, {indice.Totale} immagini nella cartella.");

        var controllo = new Controllo();
        foreach (var f in Etichetta.FontSostituiti)
            controllo.Avviso(null, null, $"font {f} non trovato in Windows: la grafica dei testi può risultare diversa");

        Console.WriteLine("Controllo dei dati...");
        var etichette = Controlla(tabella, indice, controllo);
        controllo.Stampa();

        if (controllo.Errori.Count > 0)
        {
            Console.WriteLine();
            Console.WriteLine($"PDF NON creato: correggere {(controllo.Errori.Count == 1 ? "l'errore" : $"i {controllo.Errori.Count} errori")} e rilanciare.");
            return Esito.DatiNonValidi;
        }

        int pagine = Impaginazione.Pagine(etichette.Count);
        Console.WriteLine();
        Console.WriteLine($"Controllo superato. Etichette: {etichette.Count} da {tabella.Righe.Count} righe, {pagine} pagine.");
        ScriviPdf(output, etichette, sfondo, pagine);
        Console.WriteLine($"PDF salvato in: {output}");
        if (controllo.Avvisi.Count > 0)
            Console.WriteLine($"Attenzione: ci sono {controllo.Avvisi.Count} avvisi da controllare (vedi sopra).");
        return Esito.Ok;
    }

    // ── Parametri ───────────────────────────────────────────────────────────────

    static (string? Excel, string Cartella, string Output, string? Foglio, SKColor Sfondo) LeggiParametri(string[] args)
    {
        if (args.Length == 0 || args.Any(a => a is "-h" or "--help" or "/?"))
            return (null, "", "", null, SKColors.White);

        string? output = null, foglio = null, colore = null;
        var posizionali = new List<string>();
        for (int i = 0; i < args.Length; i++)
        {
            var a = args[i];
            if (!a.StartsWith("--"))
            {
                posizionali.Add(a);
                continue;
            }
            if (a.ToLowerInvariant() is not ("--output" or "--foglio" or "--sfondo"))
                throw new ErroreGeneratore(Esito.Parametri, $"opzione sconosciuta: {a}");
            if (i + 1 >= args.Length || args[i + 1].StartsWith("--") || args[i + 1].Trim() == "")
                throw new ErroreGeneratore(Esito.Parametri, $"manca il valore dopo {a}");
            var valore = args[++i];
            switch (a.ToLowerInvariant())
            {
                case "--output": output = valore; break;
                case "--foglio": foglio = valore; break;
                case "--sfondo": colore = valore; break;
            }
        }

        if (posizionali.Count < 2)
            throw new ErroreGeneratore(Esito.Parametri,
                posizionali.Count == 0 ? "mancano il file Excel e la cartella immagini" : "manca la cartella immagini");
        if (posizionali.Count > 2)
            throw new ErroreGeneratore(Esito.Parametri, $"troppi parametri: {string.Join(" | ", posizionali)}",
                "i percorsi con spazi vanno racchiusi tra virgolette");

        string excel = Percorso(posizionali[0], "file Excel");
        string cartella = Percorso(posizionali[1], "cartella immagini");
        output = Percorso(output ?? Path.Combine(Path.GetDirectoryName(excel)!,
                                                 Path.GetFileNameWithoutExtension(excel) + " - impaginato.pdf"), "PDF");
        return (excel, cartella, output, foglio, Colore(colore ?? "giallo"));
    }

    static string Percorso(string p, string cosa)
    {
        p = p.Trim().Trim('"');
        if (p == "" || p.IndexOfAny(Path.GetInvalidPathChars()) >= 0)
            throw new ErroreGeneratore(Esito.Parametri, $"percorso del {cosa} non valido: \"{p}\"");
        try
        {
            return Path.GetFullPath(p);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            throw new ErroreGeneratore(Esito.Parametri, $"percorso del {cosa} non valido: \"{p}\" ({ex.Message})", null, ex);
        }
    }

    static SKColor Colore(string nome) => nome.Trim().ToLowerInvariant() switch
    {
        "giallo" or "arancio" or "arancione" => new SKColor(0xFF, 0xC0, 0x00),
        "bianco" => SKColors.White,
        var s when s.StartsWith('#') && s.Length == 7 && SKColor.TryParse(s, out var c) => c,
        _ => throw new ErroreGeneratore(Esito.Parametri, $"colore di sfondo non valido: \"{nome}\"",
                                        "usare giallo, bianco oppure un codice esadecimale come #FFC000"),
    };

    static void ControllaOutput(string output)
    {
        if (!string.Equals(Path.GetExtension(output), ".pdf", StringComparison.OrdinalIgnoreCase))
            throw new ErroreGeneratore(Esito.Parametri, $"il file di uscita deve avere estensione .pdf: {output}");
        if (Directory.Exists(output))
            throw new ErroreGeneratore(Esito.Parametri, $"il percorso del PDF è una cartella: {output}");
        var dir = Path.GetDirectoryName(output)!;
        if (!Directory.Exists(dir))
            throw new ErroreGeneratore(Esito.ScritturaPdf, $"la cartella del PDF non esiste: {dir}",
                "crearla oppure indicare un altro percorso con --output");
        if (File.Exists(output))
        {
            try
            {
                using var _ = File.Open(output, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (UnauthorizedAccessException ex)
            {
                throw new ErroreGeneratore(Esito.ScritturaPdf, $"il PDF esistente è di sola lettura o protetto: {output}",
                    "togliere la sola lettura o scegliere un altro nome con --output", ex);
            }
            catch (IOException ex)
            {
                throw new ErroreGeneratore(Esito.ScritturaPdf, $"il PDF è aperto in un altro programma: {output}",
                    "chiuderlo (es. Acrobat o il browser) e rilanciare", ex);
            }
        }
    }

    // ── Controllo righe ─────────────────────────────────────────────────────────

    static List<DatiEtichetta> Controlla(TabellaEtichette t, Immagini indice, Controllo k)
    {
        var cache = new Dictionary<string, ImmagineProdotto?>(StringComparer.OrdinalIgnoreCase);
        var etichette = new List<DatiEtichetta>();
        var perBarcode = new Dictionary<string, List<RigaExcel>>();
        bool haQuantita = t.Colonne.Contains(Excel.Quantita);

        foreach (var r in t.Righe)
        {
            bool rigaValida = true;
            foreach (var (col, cella) in r.Celle)
                if (cella.ErroreExcel != null)
                {
                    k.Errore(r.Numero, col, $"la cella contiene un errore di Excel ({ErroreExcel(cella.ErroreExcel)})");
                    rigaValida = false;
                }

            var immagine = ControllaImmagine(r, indice, cache, k);
            rigaValida &= immagine != null;

            string? barcode = r[Excel.Barcode];
            if (r.Celle[Excel.Barcode].ErroreExcel != null)
                barcode = null;
            else if (Ean13.Problema(barcode) is { } problema)
            {
                if (barcode == "")
                    k.Avviso(r.Numero, Excel.Barcode, "barcode vuoto: l'etichetta avrà la zona barcode bianca");
                else
                {
                    k.Errore(r.Numero, Excel.Barcode, problema);
                    rigaValida = false;
                }
                barcode = null;
            }
            else
            {
                if (!perBarcode.TryGetValue(barcode, out var stesse))
                    perBarcode[barcode] = stesse = [];
                stesse.Add(r);
            }

            if (r[Excel.Sigla] == "")
                k.Avviso(r.Numero, Excel.Sigla, "codice sigla vuoto");
            if (r[Excel.Modello] == "")
                k.Avviso(r.Numero, Excel.Modello, "modello vuoto");

            int copie = haQuantita ? ControllaQuantita(r, k) : 1;
            if (copie <= 0)
                rigaValida = false;

            if (!rigaValida)
                continue;
            var dati = new DatiEtichetta(r[Excel.Sigla], r[Excel.Modello], r[Excel.Materiale],
                                         r[Excel.Dimensioni], barcode, immagine);
            etichette.AddRange(Enumerable.Repeat(dati, copie));
        }

        foreach (var (codice, righe) in perBarcode)
        {
            var perSigla = righe.GroupBy(x => x[Excel.Sigla]).ToList();
            if (perSigla.Count < 2)
                continue;
            var dettaglio = perSigla.Select(g =>
            {
                var numeri = g.Select(x => x.Numero).ToList();
                var elenco = string.Join(", ", numeri.Take(5)) + (numeri.Count > 5 ? $" e altre {numeri.Count - 5}" : "");
                return $"{(g.Key == "" ? "(sigla vuota)" : g.Key)} alle righe {elenco}";
            });
            k.Avviso(righe[0].Numero, Excel.Barcode,
                $"barcode {codice} usato per codici sigla diversi: {string.Join("; ", dettaglio)}");
        }

        return etichette;
    }

    static ImmagineProdotto? ControllaImmagine(RigaExcel r, Immagini indice,
                                               Dictionary<string, ImmagineProdotto?> cache, Controllo k)
    {
        var valore = r[Excel.Immagine];
        var nome = Immagini.NomeFile(valore);
        if (nome == "")
        {
            k.Errore(r.Numero, Excel.Immagine, "nessuna immagine indicata");
            return null;
        }

        var ricerca = indice.Trova(valore);
        switch (ricerca.Stato)
        {
            case StatoImmagine.Mancante:
                k.Errore(r.Numero, Excel.Immagine, $"immagine {nome} non trovata nella cartella");
                return null;
            case StatoImmagine.FormatoNonSupportato:
                k.Errore(r.Numero, Excel.Immagine,
                    $"immagine {nome} trovata come {Path.GetFileName(ricerca.Percorso)}, ma il formato {Path.GetExtension(ricerca.Percorso)} non è supportato: convertirla in JPG o PNG");
                return null;
        }

        var percorso = ricerca.Percorso!;
        if (!string.Equals(Path.GetFileName(percorso), nome, StringComparison.OrdinalIgnoreCase))
            k.Avviso(r.Numero, Excel.Immagine, $"nell'Excel c'è \"{nome}\", è stata usata {Path.GetFileName(percorso)}");

        if (!cache.TryGetValue(percorso, out var img))
        {
            img = CaricaImmagine(percorso, out var errore, out var lato);
            cache[percorso] = img;
            if (img == null)
                k.Errore(r.Numero, Excel.Immagine, $"immagine {Path.GetFileName(percorso)} {errore}");
            else if (lato < LatoMinimoImmagine)
                k.Avviso(r.Numero, Excel.Immagine,
                    $"immagine {Path.GetFileName(percorso)} molto piccola ({lato} px): in stampa risulterà sgranata");
        }
        else if (img == null)
            k.Errore(r.Numero, Excel.Immagine, $"immagine {Path.GetFileName(percorso)} non leggibile (vedi riga precedente)");
        return img;
    }

    static int ControllaQuantita(RigaExcel r, Controllo k)
    {
        var testo = r[Excel.Quantita];
        if (testo == "")
        {
            k.Avviso(r.Numero, Excel.Quantita, "quantità vuota: verrà stampata 1 copia");
            return 1;
        }
        if (!double.TryParse(testo.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out var q))
        {
            k.Errore(r.Numero, Excel.Quantita, $"quantità \"{testo}\" non è un numero");
            return 0;
        }
        if (q != Math.Floor(q))
        {
            k.Errore(r.Numero, Excel.Quantita, $"quantità {testo} non è un numero intero");
            return 0;
        }
        if (q < 1)
        {
            k.Errore(r.Numero, Excel.Quantita, $"quantità {testo} non valida: deve essere almeno 1 (togliere la riga se non va stampata)");
            return 0;
        }
        if (q > QuantitaSospetta)
            k.Avviso(r.Numero, Excel.Quantita, $"quantità molto alta ({q}): verificare che sia corretta");
        return (int)q;
    }

    static string ErroreExcel(string codice) => codice switch
    {
        "NA" => "#N/D",
        "REF" => "#RIF!",
        "VALUE" => "#VALORE!",
        "DIV0" => "#DIV/0!",
        "NAME" => "#NOME?",
        "NUM" => "#NUM!",
        "NULL" => "#NULLO!",
        _ => codice,
    };

    /// <summary>
    /// Decodifica, appoggia la trasparenza su bianco (come il sorgente) e riduce l'immagine
    /// alla risoluzione di stampa (300 dpi sul riquadro), così il PDF resta leggero.
    /// </summary>
    static ImmagineProdotto? CaricaImmagine(string percorso, out string errore, out int lato)
    {
        errore = "";
        lato = 0;
        SKBitmap? originale;
        try
        {
            if (new FileInfo(percorso).Length == 0)
            {
                errore = "è un file vuoto (0 byte)";
                return null;
            }
            originale = SKBitmap.Decode(percorso);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            errore = $"non leggibile ({ex.Message})";
            return null;
        }
        if (originale == null)
        {
            errore = "danneggiata o non è una vera immagine (estensione diversa dal contenuto?)";
            return null;
        }

        using (originale)
        {
            lato = Math.Max(originale.Width, originale.Height);
            var r = Etichetta.RiquadroImmagine(originale.Width, originale.Height);
            double fattore = 300 / 190.5;
            int w = Math.Min(originale.Width, Math.Max(1, (int)Math.Round(r.Width * fattore)));
            int h = Math.Min(originale.Height, Math.Max(1, (int)Math.Round(r.Height * fattore)));

            var info = new SKImageInfo(w, h, SKColorType.Rgba8888, SKAlphaType.Opaque);
            using var surface = SKSurface.Create(info);
            surface.Canvas.Clear(SKColors.White);
            using var src = SKImage.FromBitmap(originale);
            surface.Canvas.DrawImage(src, SKRect.Create(w, h), new SKSamplingOptions(SKCubicResampler.Mitchell));
            return new ImmagineProdotto(surface.Snapshot(), r);
        }
    }

    // ── Scrittura PDF ───────────────────────────────────────────────────────────

    static void ScriviPdf(string output, List<DatiEtichetta> etichette, SKColor sfondo, int pagine)
    {
        var temporaneo = output + ".tmp";
        try
        {
            Impaginazione.CreaPdf(temporaneo, etichette, sfondo, p => Console.WriteLine($"  pagina {p}/{pagine}"));
            File.Move(temporaneo, output, overwrite: true);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            try { File.Delete(temporaneo); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            const int DiscoPieno = unchecked((int)0x80070070), DiscoPieno2 = unchecked((int)0x80070027);
            var (msg, sugg) = ex switch
            {
                UnauthorizedAccessException => ("permessi insufficienti per scrivere il PDF",
                                                 "scegliere una cartella in cui si può scrivere con --output"),
                IOException { HResult: DiscoPieno or DiscoPieno2 } => ("spazio su disco esaurito", "liberare spazio e rilanciare"),
                _ => ($"impossibile scrivere il PDF ({ex.Message})", "se il PDF è aperto in un altro programma chiuderlo e rilanciare"),
            };
            throw new ErroreGeneratore(Esito.ScritturaPdf, msg, sugg, ex);
        }
    }
}
