using System.Globalization;
using ExcelDataReader;

namespace GeneratoreEtichette;

sealed class RigaExcel(int numero, Dictionary<string, string> valori)
{
    public int Numero { get; } = numero;

    public string this[string colonna] => valori.TryGetValue(colonna, out var v) ? v : "";

    public bool Ha(string colonna) => valori.ContainsKey(colonna);
}

static class Excel
{
    public static List<RigaExcel> Leggi(string percorso, string? foglio)
    {
        using var stream = File.Open(percorso, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream);

        var nomi = new List<string>();
        do
        {
            nomi.Add(reader.Name);
            if (foglio == null || string.Equals(reader.Name, foglio, StringComparison.OrdinalIgnoreCase))
                return LeggiFoglio(reader);
        } while (reader.NextResult());

        throw new InvalidOperationException(
            $"Foglio \"{foglio}\" non trovato. Fogli presenti: {string.Join(", ", nomi)}");
    }

    static List<RigaExcel> LeggiFoglio(IExcelDataReader reader)
    {
        var righe = new List<RigaExcel>();
        if (!reader.Read())
            return righe;

        var intestazioni = new Dictionary<int, string>();
        for (int i = 0; i < reader.FieldCount; i++)
        {
            var nome = Testo(reader.GetValue(i)).ToUpperInvariant();
            if (nome != "" && !intestazioni.ContainsValue(nome))
                intestazioni[i] = nome;
        }

        int numero = 1;
        while (reader.Read())
        {
            numero++;
            var valori = new Dictionary<string, string>();
            foreach (var (i, nome) in intestazioni)
                valori[nome] = i < reader.FieldCount ? Testo(reader.GetValue(i)) : "";
            if (valori.Values.Any(v => v != ""))
                righe.Add(new RigaExcel(numero, valori));
        }
        return righe;
    }

    static string Testo(object? valore) => valore switch
    {
        null => "",
        double d when d == Math.Floor(d) && Math.Abs(d) < 1e15 => d.ToString("0", CultureInfo.InvariantCulture),
        double d => d.ToString(CultureInfo.InvariantCulture),
        IFormattable f => f.ToString(null, CultureInfo.InvariantCulture).Trim(),
        _ => valore.ToString()?.Trim() ?? "",
    };
}
