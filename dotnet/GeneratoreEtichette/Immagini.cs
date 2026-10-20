namespace GeneratoreEtichette;

enum StatoImmagine { Trovata, Mancante, FormatoNonSupportato }

sealed record RicercaImmagine(StatoImmagine Stato, string? Percorso);

/// <summary>Indicizza la cartella immagini (sottocartelle comprese) e risolve i nomi dell'Excel.</summary>
sealed class Immagini
{
    public static readonly string[] Leggibili = [".jpg", ".jpeg", ".jfif", ".png", ".bmp", ".gif", ".webp"];
    static readonly string[] NonSupportate = [".tif", ".tiff", ".ai", ".psd", ".pdf", ".eps", ".svg", ".heic"];
    static readonly string[] Estensioni = [.. Leggibili, .. NonSupportate];

    readonly Dictionary<string, string> perNome = new(StringComparer.OrdinalIgnoreCase);
    readonly Dictionary<string, string> perRadice = new(StringComparer.OrdinalIgnoreCase);
    readonly Dictionary<string, string> nonSupportatePerRadice = new(StringComparer.OrdinalIgnoreCase);

    public int Totale => perNome.Count;

    public Immagini(string cartella)
    {
        if (File.Exists(cartella))
            throw new ErroreGeneratore(Esito.Parametri, $"il percorso immagini indicato è un file, non una cartella: {cartella}",
                "come secondo parametro va indicata la cartella che contiene le immagini");
        if (!Directory.Exists(cartella))
            throw new ErroreGeneratore(Esito.CartellaImmagini, $"cartella immagini non trovata: {cartella}",
                "controllare il percorso; se contiene spazi va racchiuso tra virgolette");

        List<string> file;
        try
        {
            var opzioni = new EnumerationOptions { RecurseSubdirectories = true, IgnoreInaccessible = true };
            file = Directory.EnumerateFiles(cartella, "*", opzioni)
                .Where(p => !p.Split(Path.DirectorySeparatorChar).Contains("__MACOSX"))
                .Where(p => !Path.GetFileName(p).StartsWith('.'))
                .OrderBy(p => p.Count(c => c == Path.DirectorySeparatorChar))
                .ThenBy(p => p, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new ErroreGeneratore(Esito.CartellaImmagini, "accesso negato alla cartella immagini",
                "controllare i permessi o copiare le immagini in una cartella locale", ex);
        }
        catch (IOException ex)
        {
            throw new ErroreGeneratore(Esito.CartellaImmagini, $"impossibile leggere la cartella immagini ({ex.Message})",
                "se è una cartella di rete verificare che sia raggiungibile", ex);
        }

        foreach (var p in file)
        {
            var ext = Path.GetExtension(p).ToLowerInvariant();
            if (Leggibili.Contains(ext))
            {
                perNome.TryAdd(Path.GetFileName(p), p);
                perRadice.TryAdd(Radice(Path.GetFileName(p)), p);
            }
            else if (NonSupportate.Contains(ext))
                nonSupportatePerRadice.TryAdd(Radice(Path.GetFileName(p)), p);
        }

        if (perNome.Count == 0)
            throw new ErroreGeneratore(Esito.CartellaImmagini,
                $"la cartella immagini non contiene immagini utilizzabili ({string.Join(", ", Leggibili)}): {cartella}",
                nonSupportatePerRadice.Count > 0
                    ? $"ci sono {nonSupportatePerRadice.Count} file in formati non supportati: convertirli in JPG o PNG"
                    : "verificare di aver indicato la cartella giusta");
    }

    /// <summary>
    /// Nome esatto, poi nome senza estensione: copre estensioni mancanti, .ai al posto
    /// di .jpg, .jpg.jpg e percorsi UNC senza separatore prima del nome file.
    /// </summary>
    public RicercaImmagine Trova(string immagine)
    {
        var nome = NomeFile(immagine);
        if (perNome.TryGetValue(nome, out var esatto))
            return new(StatoImmagine.Trovata, esatto);

        var radici = new List<string> { Radice(nome) };
        const string prefisso = "IMMAGINI_ARTICOLI_";
        if (radici[0].StartsWith(prefisso, StringComparison.OrdinalIgnoreCase))
            radici.Add(radici[0][prefisso.Length..]);

        foreach (var r in radici)
            if (perRadice.TryGetValue(r, out var p))
                return new(StatoImmagine.Trovata, p);
        foreach (var r in radici)
            if (nonSupportatePerRadice.TryGetValue(r, out var p))
                return new(StatoImmagine.FormatoNonSupportato, p);
        return new(StatoImmagine.Mancante, null);
    }

    public static string NomeFile(string immagine) => immagine.Replace('\\', '/').Split('/')[^1].Trim();

    static string Radice(string nome)
    {
        nome = nome.Trim();
        while (Estensioni.Contains(Path.GetExtension(nome).ToLowerInvariant()))
            nome = Path.GetFileNameWithoutExtension(nome);
        return nome.Trim();
    }
}
