namespace GeneratoreEtichette;

/// <summary>Codici di uscita del programma (documentati in LEGGIMI.txt).</summary>
enum Esito
{
    Ok = 0,
    Parametri = 1,
    Excel = 2,
    CartellaImmagini = 3,
    DatiNonValidi = 4,
    ScritturaPdf = 5,
    Imprevisto = 9,
}

/// <summary>Errore previsto: interrompe il programma con un messaggio chiaro e un suggerimento.</summary>
sealed class ErroreGeneratore(Esito esito, string messaggio, string? suggerimento = null, Exception? causa = null)
    : Exception(messaggio, causa)
{
    public Esito Esito { get; } = esito;
    public string? Suggerimento { get; } = suggerimento;
}

sealed record Segnalazione(int? Riga, string? Colonna, string Testo)
{
    public override string ToString()
    {
        var dove = (Riga, Colonna) switch
        {
            (int r, string c) => $"Riga {r} [{c}]: ",
            (int r, null) => $"Riga {r}: ",
            (null, string c) => $"[{c}]: ",
            _ => "",
        };
        return dove + Testo;
    }
}

/// <summary>Raccoglie tutti i problemi trovati nei dati, così l'utente li vede tutti insieme.</summary>
sealed class Controllo
{
    public List<Segnalazione> Errori { get; } = [];
    public List<Segnalazione> Avvisi { get; } = [];

    public void Errore(int? riga, string? colonna, string testo) => Errori.Add(new(riga, colonna, testo));
    public void Avviso(int? riga, string? colonna, string testo) => Avvisi.Add(new(riga, colonna, testo));

    public void Stampa()
    {
        if (Errori.Count > 0)
        {
            Console.WriteLine();
            Console.WriteLine($"ERRORI ({Errori.Count}) - il PDF NON verrà creato:");
            foreach (var e in Errori)
                Console.WriteLine($"  {e}");
        }
        if (Avvisi.Count > 0)
        {
            Console.WriteLine();
            Console.WriteLine($"AVVISI ({Avvisi.Count}) - non bloccano il PDF, ma vanno controllati:");
            foreach (var a in Avvisi)
                Console.WriteLine($"  {a}");
        }
    }
}
