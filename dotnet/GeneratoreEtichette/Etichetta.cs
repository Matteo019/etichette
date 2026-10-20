using SkiaSharp;

namespace GeneratoreEtichette;

sealed record ImmagineProdotto(SKImage Immagine, SKRectI Riquadro);

sealed record DatiEtichetta(string Sigla, string Modello, string Materiale, string Dimensioni,
                            string? Barcode, ImmagineProdotto? Immagine);

/// <summary>
/// Porting di genera_label() di codice/genera_etichette.py. Si disegna in "pixel etichetta":
/// 750×375 a 190.5 DPI = 100×50 mm, le stesse unità del sorgente Python.
/// </summary>
static class Etichetta
{
    public const int W = 750, H = 375;
    const int CenterX = (int)(W * 0.731);
    const int MaxTxtW = (int)(W * 0.385);
    const int LeftBoxW = (int)(W * 0.48);
    const int BoxW = LeftBoxW - 20, BoxH = H - 20;

    // Opzioni python-barcode usate dal sorgente (ImageWriter a 300 dpi incollato 1:1).
    const double BcDpi = 300, Modulo = 0.25, AltezzaBarre = 6.0, FattoreGuardia = 1.2;
    const double Quiet = 2.0, MargineSup = 1.0, FontPt = 7;

    static readonly SKTypeface Grassetto = Carattere("arialbd.ttf", SKFontStyle.Bold);
    static readonly SKTypeface Normale = Carattere("arial.ttf", SKFontStyle.Normal);

    /// <summary>Font Arial non trovati: si usa il font di sistema più simile (la grafica può cambiare).</summary>
    public static List<string> FontSostituiti { get; } = [];

    static SKTypeface Carattere(string file, SKFontStyle stile)
    {
        var accanto = Path.Combine(AppContext.BaseDirectory, file);
        var sistema = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), file);
        foreach (var p in new[] { accanto, sistema })
            if (File.Exists(p) && SKTypeface.FromFile(p) is { } tf)
                return tf;
        FontSostituiti.Add(file);
        return SKTypeface.FromFamilyName("Arial", stile) ?? SKTypeface.Default;
    }

    /// <summary>Riquadro (in pixel etichetta) in cui finisce l'immagine, come ImageOps.pad.</summary>
    public static SKRectI RiquadroImmagine(int larghezza, int altezza)
    {
        double ratio = (double)larghezza / altezza;
        int w, h;
        if (ratio > (double)BoxW / BoxH)
        {
            w = BoxW;
            h = (int)Math.Round(BoxW / ratio);
        }
        else
        {
            h = BoxH;
            w = (int)Math.Round(BoxH * ratio);
        }
        int x = (LeftBoxW - BoxW) / 2 + (int)Math.Round((BoxW - w) * 0.5);
        int y = (H - BoxH) / 2 + (int)Math.Round((BoxH - h) * 0.5);
        return SKRectI.Create(x, y, w, h);
    }

    public static void Disegna(SKCanvas c, DatiEtichetta e, SKColor sfondo)
    {
        using (var fondo = new SKPaint { Color = sfondo })
            c.DrawRect(0, 0, W, H, fondo);

        if (e.Immagine is { } img)
            c.DrawImage(img.Immagine, img.Riquadro, new SKSamplingOptions(SKCubicResampler.Mitchell));

        if (e.Barcode is { } codice)
            DisegnaBarcode(c, codice);

        Testo(c, e.Sigla, 0.10, 0.390);
        Testo(c, e.Modello, 0.25, 0.505, maxWPct: 0.44, allarga: 1.12);
        Testo(c, e.Materiale, 0.12, 0.755, allarga: 1.06);
        Testo(c, e.Dimensioni, 0.11, 0.880, allarga: 1.06);
    }

    /// <summary>Riga centrata: ritagliata sull'inchiostro, allargata, stretta se sfora (draw_text).</summary>
    static void Testo(SKCanvas c, string testo, double hPct, double yPct,
                      double? maxWPct = null, double allarga = 1.0)
    {
        if (testo == "")
            return;
        using var font = new SKFont(Grassetto, (int)(H * hPct)) { Subpixel = true };
        font.MeasureText(testo, out var ink);
        if (ink.Width <= 0)
            return;
        int maxW = maxWPct is { } m ? (int)(W * m) : MaxTxtW;
        int nw = Math.Min(Math.Max(1, (int)(Math.Ceiling(ink.Width) * allarga)), maxW);
        using var paint = new SKPaint { Color = SKColors.Black, IsAntialias = true };
        c.Save();
        c.Translate(CenterX - nw / 2, (int)(H * yPct));
        c.Scale(nw / ink.Width, 1);
        c.DrawText(testo, -ink.Left, -ink.Top, font, paint);
        c.Restore();
    }

    static double Px(double mm) => mm * BcDpi / 25.4;

    /// <summary>EAN-13 con barre di guardia, stessa geometria e arrotondamenti di python-barcode.</summary>
    static void DisegnaBarcode(SKCanvas c, string codice)
    {
        string moduli = Ean13.Moduli(codice);
        int larghezza = (int)Px(2 * Quiet + moduli.Length * Modulo);
        float ox = CenterX - larghezza / 2, oy = (int)(H * 0.056);

        using var nero = new SKPaint { Color = SKColors.Black, IsAntialias = false };
        var inizioGuardie = new List<double>();
        var fineGuardie = new List<double>();
        bool inGuardia = false;
        double x = Quiet;
        double fattore = 1;

        for (int i = 0; i < moduli.Length;)
        {
            int j = i;
            while (j < moduli.Length && moduli[j] == moduli[i])
                j++;
            char ch = moduli[i];
            double w = Modulo * (j - i);
            fattore = ch == '1' ? 1 : FattoreGuardia;
            if (ch != '0')
            {
                if (inGuardia && fattore == 1) { fineGuardie.Add(x); inGuardia = false; }
                else if (!inGuardia && fattore != 1) { inizioGuardie.Add(x); inGuardia = true; }

                double h = AltezzaBarre * fattore;
                float x0 = (float)Math.Floor(Px(x)), x1 = (float)Math.Floor(Px(x + w));
                float y0 = (float)Math.Floor(Px(MargineSup)), y1 = (float)Math.Floor(Px(MargineSup + h)) + 1;
                c.DrawRect(ox + x0, oy + y0, x1 - x0, y1 - y0, nero);
            }
            x += w;
            i = j;
        }
        if (fattore != 1)
            fineGuardie.Add(x);

        double inizio = Quiet;
        inizioGuardie.RemoveAt(0);
        var posizioni = new List<double> { inizio - 4 * Modulo };
        for (int k = 0; k < Math.Min(inizioGuardie.Count, fineGuardie.Count); k++)
            posizioni.Add(fineGuardie[k] + (inizioGuardie[k] - fineGuardie[k]) / 2);

        string[] blocchi = [codice[..1], codice[1..7], codice[7..]];
        double yTesto = MargineSup + AltezzaBarre + FontPt * 0.352777778;
        using var font = new SKFont(Normale, (int)Px(FontPt * 0.352777778)) { Subpixel = true };
        using var testo = new SKPaint { Color = SKColors.Black, IsAntialias = true };
        float baseline = (float)Px(yTesto) - font.Metrics.Descent;
        for (int k = 0; k < blocchi.Length; k++)
            c.DrawText(blocchi[k], ox + (float)Px(posizioni[k]), oy + baseline,
                       SKTextAlign.Center, font, testo);
    }
}

static class Ean13
{
    static readonly string[] A = ["0001101", "0011001", "0010011", "0111101", "0100011",
                                  "0110001", "0101111", "0111011", "0110111", "0001011"];
    static readonly string[] B = ["0100111", "0110011", "0011011", "0100001", "0011101",
                                  "0111001", "0000101", "0010001", "0001001", "0010111"];
    static readonly string[] C = ["1110010", "1100110", "1101100", "1000010", "1011100",
                                  "1001110", "1010000", "1000100", "1001000", "1110100"];
    static readonly string[] Parita = ["AAAAAA", "AABABB", "AABBAB", "AABBBA", "ABAABB",
                                       "ABBAAB", "ABBBAA", "ABABAB", "ABABBA", "ABBABA"];

    /// <summary>Null se il codice è un EAN-13 valido, altrimenti il motivo.</summary>
    public static string? Problema(string codice)
    {
        if (codice == "")
            return "barcode vuoto";
        if (codice.Contains('E', StringComparison.OrdinalIgnoreCase) && codice.Contains('+'))
            return $"barcode \"{codice}\" in notazione scientifica: formattare la colonna come Testo o Numero senza decimali";
        var estranei = codice.Where(ch => !char.IsAsciiDigit(ch)).Distinct().ToArray();
        if (estranei.Length > 0)
            return $"barcode \"{codice}\" contiene caratteri non numerici ({string.Join(" ", estranei.Select(ch => ch == ' ' ? "spazio" : $"'{ch}'"))})";
        if (codice.Length != 13)
            return $"barcode \"{codice}\" ha {codice.Length} cifre invece di 13";
        int attesa = CifraControllo(codice);
        if (attesa != codice[12] - '0')
            return $"barcode \"{codice}\" ha la cifra di controllo errata: l'ultima cifra dovrebbe essere {attesa}";
        return null;
    }

    static int CifraControllo(string codice)
    {
        int somma = 0;
        for (int i = 0; i < 12; i++)
            somma += (codice[i] - '0') * (i % 2 == 0 ? 1 : 3);
        return (10 - somma % 10) % 10;
    }

    /// <summary>95 moduli: '1' barra, '0' spazio, 'G' barra di guardia.</summary>
    public static string Moduli(string codice)
    {
        var p = Parita[codice[0] - '0'];
        var s = new System.Text.StringBuilder("G0G");
        for (int i = 0; i < 6; i++)
            s.Append((p[i] == 'A' ? A : B)[codice[i + 1] - '0']);
        s.Append("0G0G0");
        for (int i = 7; i < 13; i++)
            s.Append(C[codice[i] - '0']);
        s.Append("G0G");
        return s.ToString();
    }
}
