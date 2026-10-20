namespace GeneratoreEtichette;

/// <summary>Indicizza la cartella immagini (sottocartelle comprese) e risolve i nomi dell'Excel.</summary>
sealed class Immagini
{
    static readonly string[] Leggibili = [".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp"];
    static readonly string[] Ignorate = [.. Leggibili, ".tif", ".tiff", ".ai", ".psd", ".pdf", ".eps"];

    readonly Dictionary<string, string> perNome = new(StringComparer.OrdinalIgnoreCase);
    readonly Dictionary<string, string> perRadice = new(StringComparer.OrdinalIgnoreCase);

    public Immagini(string cartella)
    {
        var file = Directory.EnumerateFiles(cartella, "*", SearchOption.AllDirectories)
            .Where(p => !p.Split(Path.DirectorySeparatorChar).Any(parte => parte == "__MACOSX"))
            .Where(p => !Path.GetFileName(p).StartsWith('.'))
            .Where(p => Leggibili.Contains(Path.GetExtension(p).ToLowerInvariant()))
            .OrderBy(p => p.Count(c => c == Path.DirectorySeparatorChar))
            .ThenBy(p => p, StringComparer.OrdinalIgnoreCase);

        foreach (var p in file)
        {
            perNome.TryAdd(Path.GetFileName(p), p);
            perRadice.TryAdd(Radice(Path.GetFileName(p)), p);
        }
    }

    /// <summary>
    /// Nome esatto, poi nome senza estensione: copre estensioni mancanti, .ai al posto
    /// di .jpg, .jpg.jpg e percorsi UNC senza separatore prima del nome file.
    /// </summary>
    public string? Trova(string immagine)
    {
        var nome = immagine.Replace('\\', '/').Split('/')[^1].Trim();
        if (nome == "")
            return null;
        if (perNome.TryGetValue(nome, out var esatto))
            return esatto;

        var radice = Radice(nome);
        if (perRadice.TryGetValue(radice, out var p))
            return p;
        const string prefisso = "IMMAGINI_ARTICOLI_";
        if (radice.StartsWith(prefisso, StringComparison.OrdinalIgnoreCase)
            && perRadice.TryGetValue(radice[prefisso.Length..], out p))
            return p;
        return null;
    }

    public static string NomeFile(string immagine) => immagine.Replace('\\', '/').Split('/')[^1].Trim();

    static string Radice(string nome)
    {
        nome = nome.Trim();
        while (Ignorate.Contains(Path.GetExtension(nome).ToLowerInvariant()))
            nome = Path.GetFileNameWithoutExtension(nome);
        return nome.Trim();
    }
}
