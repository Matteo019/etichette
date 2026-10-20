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
          --foglio <nome>       foglio Excel da leggere (default: il primo)
          --sfondo <colore>     giallo (default), bianco oppure #RRGGBB

        Prima di creare il PDF controlla che ogni riga abbia la sua immagine:
        se ne manca anche una sola elenca le righe e il PDF NON viene creato.
        """;

    static int Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);   // richiesto da ExcelDataReader
        try
        {
            return Esegui(args);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"ERRORE: {ex.Message}");
            return 1;
        }
    }

    static int Esegui(string[] args)
    {
        string? output = null, foglio = null, coloreSfondo = "giallo";
        var posizionali = new List<string>();
        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i].ToLowerInvariant())
            {
                case "--output" when i + 1 < args.Length: output = args[++i]; break;
                case "--foglio" when i + 1 < args.Length: foglio = args[++i]; break;
                case "--sfondo" when i + 1 < args.Length: coloreSfondo = args[++i]; break;
                case "-h" or "--help" or "/?": Console.WriteLine(Uso); return 0;
                default: posizionali.Add(args[i]); break;
            }
        }
        if (posizionali.Count != 2)
        {
            Console.WriteLine(Uso);
            return 1;
        }

        string excel = Path.GetFullPath(posizionali[0]);
        string cartella = Path.GetFullPath(posizionali[1]);
        if (!File.Exists(excel))
            throw new FileNotFoundException($"file Excel non trovato: {excel}");
        if (!Directory.Exists(cartella))
            throw new DirectoryNotFoundException($"cartella immagini non trovata: {cartella}");
        output = Path.GetFullPath(output ?? Path.Combine(Path.GetDirectoryName(excel)!,
                                  Path.GetFileNameWithoutExtension(excel) + " - impaginato.pdf"));
        var sfondo = Colore(coloreSfondo);

        Console.WriteLine($"Excel:    {excel}");
        Console.WriteLine($"Immagini: {cartella}");
        var righe = Excel.Leggi(excel, foglio);
        Console.WriteLine($"Righe lette: {righe.Count}");
        if (righe.Count == 0)
            throw new InvalidOperationException("nessuna riga da elaborare");
        if (!righe[0].Ha("IMMAGINE"))
            throw new InvalidOperationException("colonna IMMAGINE assente nell'Excel");

        // ── Controllo immagini: blocca la generazione se ne manca anche una ──
        Console.WriteLine("Controllo immagini...");
        var indice = new Immagini(cartella);
        var cache = new Dictionary<string, ImmagineProdotto>(StringComparer.OrdinalIgnoreCase);
        var errori = new List<string>();
        var immaginiRiga = new Dictionary<RigaExcel, ImmagineProdotto>();
        foreach (var r in righe)
        {
            var nome = Immagini.NomeFile(r["IMMAGINE"]);
            if (nome == "")
            {
                errori.Add($"Riga {r.Numero}: nessuna immagine indicata nella colonna IMMAGINE");
                continue;
            }
            var percorso = indice.Trova(r["IMMAGINE"]);
            if (percorso == null)
            {
                errori.Add($"Riga {r.Numero}: immagine {nome} non trovata");
                continue;
            }
            if (!cache.TryGetValue(percorso, out var img))
            {
                img = CaricaImmagine(percorso);
                if (img == null)
                {
                    errori.Add($"Riga {r.Numero}: immagine {Path.GetFileName(percorso)} non leggibile");
                    continue;
                }
                cache[percorso] = img;
            }
            immaginiRiga[r] = img;
        }

        if (errori.Count > 0)
        {
            Console.WriteLine();
            Console.WriteLine($"IMMAGINI MANCANTI: {errori.Count} righe senza immagine.");
            foreach (var e in errori)
                Console.WriteLine($"  {e}");
            Console.WriteLine();
            Console.WriteLine("PDF NON creato. Aggiungere le immagini nella cartella e rilanciare.");
            return 2;
        }
        Console.WriteLine($"OK: tutte le {righe.Count} righe hanno l'immagine.");

        // ── Etichette ─────────────────────────────────────────────────────────
        var etichette = new List<DatiEtichetta>();
        foreach (var r in righe)
        {
            var codice = r["BARCODE"];
            string? barcode = codice;
            if (!Ean13.Valido(codice))
            {
                Console.WriteLine($"ATTENZIONE (Riga {r.Numero}): barcode non valido ({(codice == "" ? "vuoto" : codice)}) -> zona bianca");
                barcode = null;
            }
            var dati = new DatiEtichetta(r["CODICE SIGLA"], r["MODELLO"], r["MATERIALE"], r["DIMENSIONI"],
                                         barcode, immaginiRiga[r]);
            etichette.AddRange(Enumerable.Repeat(dati, Quantita(r)));
        }

        int pagine = Impaginazione.Pagine(etichette.Count);
        Console.WriteLine($"Etichette: {etichette.Count} da {righe.Count} righe, {pagine} pagine.");
        Impaginazione.CreaPdf(output, etichette, sfondo,
                              p => Console.WriteLine($"  pagina {p}/{pagine}"));
        Console.WriteLine($"PDF salvato in: {output}");
        return 0;
    }

    static int Quantita(RigaExcel r)
    {
        foreach (var col in new[] { "QUANTITA", "Q.TA", "QTA", "QUANTITÀ" })
            if (r.Ha(col))
                return double.TryParse(r[col], System.Globalization.NumberStyles.Float,
                                       System.Globalization.CultureInfo.InvariantCulture, out var q)
                    ? Math.Max((int)q, 1) : 1;
        return 1;
    }

    static SKColor Colore(string nome) => nome.ToLowerInvariant() switch
    {
        "giallo" or "arancio" => new SKColor(0xFF, 0xC0, 0x00),
        "bianco" => SKColors.White,
        _ when SKColor.TryParse(nome, out var c) => c,
        _ => throw new ArgumentException($"colore di sfondo non valido: {nome}"),
    };

    /// <summary>
    /// Decodifica, appoggia la trasparenza su bianco (come il sorgente) e riduce l'immagine
    /// alla risoluzione di stampa (300 dpi sul riquadro), così il PDF resta leggero.
    /// </summary>
    static ImmagineProdotto? CaricaImmagine(string percorso)
    {
        using var originale = SKBitmap.Decode(percorso);
        if (originale == null)
            return null;
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
