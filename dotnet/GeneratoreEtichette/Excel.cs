using System.Globalization;
using System.Text;
using ExcelDataReader;
using ExcelDataReader.Exceptions;

namespace GeneratoreEtichette;

sealed record Cella(string Testo, string? ErroreExcel);

sealed class RigaExcel(int numero, Dictionary<string, Cella> celle)
{
    public int Numero { get; } = numero;
    public Dictionary<string, Cella> Celle { get; } = celle;

    public string this[string colonna] => Celle.TryGetValue(colonna, out var c) ? c.Testo : "";
}

sealed record TabellaEtichette(string Foglio, IReadOnlyList<string> Colonne, List<RigaExcel> Righe);

static class Excel
{
    public const string Barcode = "BARCODE", Sigla = "CODICE SIGLA", Modello = "MODELLO",
                        Materiale = "MATERIALE", Dimensioni = "DIMENSIONI", Immagine = "IMMAGINE",
                        Quantita = "QUANTITA";

    public static readonly string[] Obbligatorie = [Barcode, Sigla, Modello, Immagine];
    static readonly string[] Facoltative = [Materiale, Dimensioni, Quantita];
    static readonly string[] Estensioni = [".xlsx", ".xlsm", ".xlsb", ".xls"];

    static readonly Dictionary<string, string> Alias = new()
    {
        ["Q.TA"] = Quantita, ["QTA"] = Quantita, ["Q.TA'"] = Quantita, ["QUANTITA'"] = Quantita,
        ["CODICESIGLA"] = Sigla, ["EAN"] = Barcode, ["EAN13"] = Barcode, ["EAN 13"] = Barcode,
    };

    sealed record Foglio(string Nome, List<Cella[]> Righe);

    public static void ControllaFile(string percorso)
    {
        if (Directory.Exists(percorso))
            throw new ErroreGeneratore(Esito.Parametri, $"il percorso Excel indicato è una cartella: {percorso}",
                "come primo parametro va indicato il file .xlsx, come secondo la cartella delle immagini");
        if (!File.Exists(percorso))
            throw new ErroreGeneratore(Esito.Parametri, $"file Excel non trovato: {percorso}",
                "controllare il percorso; se contiene spazi va racchiuso tra virgolette");
        var ext = Path.GetExtension(percorso).ToLowerInvariant();
        if (ext == ".csv")
            throw new ErroreGeneratore(Esito.Excel, "il file è un CSV, non un Excel",
                "aprirlo in Excel e salvarlo come \"Cartella di lavoro di Excel (.xlsx)\"");
        if (!Estensioni.Contains(ext))
            throw new ErroreGeneratore(Esito.Excel, $"estensione \"{ext}\" non supportata: servono file .xlsx, .xlsm, .xlsb o .xls");
        if (new FileInfo(percorso).Length == 0)
            throw new ErroreGeneratore(Esito.Excel, "il file Excel è vuoto (0 byte)", "il file potrebbe essere stato copiato male");
        if (Path.GetFileName(percorso).StartsWith("~$"))
            throw new ErroreGeneratore(Esito.Excel, "il file indicato è un file temporaneo di Excel (~$...)",
                "indicare il file Excel vero, senza ~$ davanti al nome");
    }

    public static TabellaEtichette Leggi(string percorso, string? nomeFoglio)
    {
        var fogli = LeggiFogli(percorso);
        if (fogli.Count == 0)
            throw new ErroreGeneratore(Esito.Excel, "il file Excel non contiene fogli");

        if (nomeFoglio != null)
        {
            var f = fogli.FirstOrDefault(x => string.Equals(x.Nome.Trim(), nomeFoglio.Trim(), StringComparison.OrdinalIgnoreCase))
                ?? throw new ErroreGeneratore(Esito.Excel, $"foglio \"{nomeFoglio}\" non trovato",
                       $"fogli presenti: {string.Join(", ", fogli.Select(x => $"\"{x.Nome}\""))}");
            return Analizza(f, fogli);
        }

        // Senza --foglio: il primo foglio che ha le colonne obbligatorie.
        var adatto = fogli.FirstOrDefault(f => f.Righe.Count > 0 && Mancanti(Intestazioni(f.Righe[0])).Count == 0);
        return Analizza(adatto ?? fogli[0], fogli);
    }

    static List<Foglio> LeggiFogli(string percorso)
    {
        try
        {
            using var stream = File.Open(percorso, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var fogli = new List<Foglio>();
            do
            {
                var righe = new List<Cella[]>();
                while (reader.Read())
                {
                    var r = new Cella[reader.FieldCount];
                    for (int i = 0; i < reader.FieldCount; i++)
                        r[i] = new Cella(Testo(reader.GetValue(i)), reader.GetCellError(i)?.ToString());
                    righe.Add(r);
                }
                fogli.Add(new Foglio(reader.Name ?? $"Foglio{fogli.Count + 1}", righe));
            } while (reader.NextResult());
            return fogli;
        }
        catch (InvalidPasswordException ex)
        {
            throw new ErroreGeneratore(Esito.Excel, "il file Excel è protetto da password",
                "aprirlo in Excel, togliere la password (File > Informazioni > Proteggi cartella di lavoro) e salvarlo", ex);
        }
        catch (HeaderException ex)
        {
            throw new ErroreGeneratore(Esito.Excel, "il file non è un Excel valido (formato non riconosciuto)",
                "potrebbe essere un file rinominato o danneggiato: aprirlo in Excel e salvarlo di nuovo come .xlsx", ex);
        }
        catch (Exception ex) when (ex is ExcelReaderException or InvalidDataException or ArgumentException
                                   or System.Xml.XmlException or EndOfStreamException)
        {
            throw new ErroreGeneratore(Esito.Excel, $"il file Excel è danneggiato o in un formato non leggibile ({ex.Message})",
                "aprirlo in Excel e salvarlo di nuovo come .xlsx", ex);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new ErroreGeneratore(Esito.Excel, "accesso negato al file Excel",
                "controllare i permessi del file o copiarlo in una cartella locale", ex);
        }
        catch (IOException ex)
        {
            throw new ErroreGeneratore(Esito.Excel, $"impossibile leggere il file Excel ({ex.Message})",
                "se è aperto in un altro programma chiuderlo e riprovare", ex);
        }
    }

    static TabellaEtichette Analizza(Foglio f, List<Foglio> tutti)
    {
        if (f.Righe.All(r => r.All(c => c.Testo == "")))
            throw new ErroreGeneratore(Esito.Excel, $"il foglio \"{f.Nome}\" è vuoto", AltriFogli(f, tutti));

        var intestazioni = Intestazioni(f.Righe[0]);
        var mancanti = Mancanti(intestazioni);
        if (mancanti.Count > 0)
        {
            var trovate = intestazioni.Values.Where(v => v != "").Distinct().ToList();
            string? suggerimento = null;
            for (int r = 1; r < Math.Min(f.Righe.Count, 15) && suggerimento == null; r++)
                if (Mancanti(Intestazioni(f.Righe[r])).Count < Obbligatorie.Length - 1)
                    suggerimento = $"le intestazioni sembrano essere alla riga {r + 1}: devono stare nella riga 1 (eliminare le righe sopra)";
            suggerimento ??= AltriFogli(f, tutti);
            suggerimento ??= $"colonne richieste: {string.Join(", ", Obbligatorie)}; facoltative: {string.Join(", ", Facoltative)}";
            throw new ErroreGeneratore(Esito.Excel,
                $"Excel non adatto: nel foglio \"{f.Nome}\" mancano le colonne {string.Join(", ", mancanti)}"
                + $" (intestazioni trovate in riga 1: {(trovate.Count == 0 ? "nessuna" : string.Join(", ", trovate))})",
                suggerimento);
        }

        var doppie = intestazioni.Values.Where(v => v != "").GroupBy(v => v).Where(g => g.Count() > 1).Select(g => g.Key).ToList();
        if (doppie.Count > 0)
            throw new ErroreGeneratore(Esito.Excel,
                $"colonne ripetute nella riga 1: {string.Join(", ", doppie)}",
                "ogni intestazione deve comparire una sola volta");

        var righe = new List<RigaExcel>();
        for (int r = 1; r < f.Righe.Count; r++)
        {
            var celle = new Dictionary<string, Cella>();
            foreach (var (i, nome) in intestazioni)
                if (nome != "")
                    celle[nome] = i < f.Righe[r].Length ? f.Righe[r][i] : new Cella("", null);
            if (celle.Values.Any(c => c.Testo != "" || c.ErroreExcel != null))
                righe.Add(new RigaExcel(r + 1, celle));
        }
        if (righe.Count == 0)
            throw new ErroreGeneratore(Esito.Excel, $"il foglio \"{f.Nome}\" ha le intestazioni ma nessuna riga di dati");

        var colonne = intestazioni.Values.Where(v => v != "").ToList();
        return new TabellaEtichette(f.Nome, colonne, righe);
    }

    static string? AltriFogli(Foglio f, List<Foglio> tutti)
    {
        var altri = tutti.Where(x => x != f && x.Righe.Count > 0 && Mancanti(Intestazioni(x.Righe[0])).Count == 0)
                         .Select(x => $"\"{x.Nome}\"").ToList();
        return altri.Count > 0 ? $"le colonne giuste sono nel foglio {string.Join(", ", altri)}: usare --foglio" : null;
    }

    static Dictionary<int, string> Intestazioni(Cella[] riga)
    {
        var d = new Dictionary<int, string>();
        for (int i = 0; i < riga.Length; i++)
            d[i] = Normalizza(riga[i].Testo);
        return d;
    }

    static List<string> Mancanti(Dictionary<int, string> intestazioni) =>
        Obbligatorie.Where(o => !intestazioni.ContainsValue(o)).ToList();

    /// <summary>Maiuscolo, senza accenti, spazi singoli; applica gli alias (Q.TA → QUANTITA...).</summary>
    static string Normalizza(string s)
    {
        var senzaAccenti = new StringBuilder();
        foreach (var ch in s.Normalize(NormalizationForm.FormD))
            if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
                senzaAccenti.Append(ch);
        var n = string.Join(' ', senzaAccenti.ToString().ToUpperInvariant().Replace('_', ' ')
                                  .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        return Alias.TryGetValue(n, out var a) ? a : n;
    }

    static string Testo(object? valore) => valore switch
    {
        null => "",
        double d when d == Math.Floor(d) && Math.Abs(d) < 1e15 => d.ToString("0", CultureInfo.InvariantCulture),
        double d => d.ToString(CultureInfo.InvariantCulture),
        DateTime dt => dt.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture),
        IFormattable f => f.ToString(null, CultureInfo.InvariantCulture).Trim(),
        _ => valore.ToString()?.Trim() ?? "",
    };
}
