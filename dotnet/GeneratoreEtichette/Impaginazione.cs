using SkiaSharp;

namespace GeneratoreEtichette;

/// <summary>
/// Foglio della fresa 970×410 mm: 19×4 etichette ruotate di 90°, margini 10/5 mm,
/// cerchi di riferimento Ø6 mm e squadretta a L (come run() del sorgente Python).
/// </summary>
static class Impaginazione
{
    const int Colonne = 19, Righe = 4, PerPagina = Colonne * Righe;
    const double Dpi = Etichetta.W / (100.0 / 25.4);   // 190.5
    const double MargineX = 10.0 / 25.4 * Dpi;          // 75 px
    const double MargineY = 5.0 / 25.4 * Dpi;           // 37.5 px
    const int PaginaW = (int)(Colonne * Etichetta.H + 2 * MargineX);
    const int PaginaH = (int)(Righe * Etichetta.W + 2 * MargineY);

    public static int Pagine(int etichette) => (etichette + PerPagina - 1) / PerPagina;

    public static void CreaPdf(string percorso, IReadOnlyList<DatiEtichetta> etichette, SKColor sfondo,
                               Action<int>? avanzamento = null)
    {
        float scala = (float)(72 / Dpi);
        var metadati = new SKDocumentPdfMetadata
        {
            Title = Path.GetFileNameWithoutExtension(percorso),
            Creator = "GeneratoreEtichette",
            RasterDpi = 300,
            EncodingQuality = 92,
        };

        using (var file = File.Create(percorso))
        using (var doc = SKDocument.CreatePdf(file, metadati))
        {
            for (int inizio = 0, pagina = 1; inizio < etichette.Count; inizio += PerPagina, pagina++)
            {
                var c = doc.BeginPage(PaginaW * scala, PaginaH * scala);
                c.Scale(scala);
                DisegnaPagina(c, etichette.Skip(inizio).Take(PerPagina).ToList(), sfondo);
                doc.EndPage();
                avanzamento?.Invoke(pagina);
            }
            doc.Close();
        }
    }

    static void DisegnaPagina(SKCanvas c, List<DatiEtichetta> etichette, SKColor sfondo)
    {
        c.Clear(SKColors.White);
        for (int j = 0; j < etichette.Count; j++)
        {
            int x = (int)(MargineX + j % Colonne * Etichetta.H);
            int y = (int)(MargineY + j / Colonne * Etichetta.W);
            c.Save();
            c.ClipRect(SKRect.Create(x, y, Etichetta.H, Etichetta.W));
            c.Translate(x, y + Etichetta.W);   // rotazione antioraria di 90° (PIL ROTATE_90)
            c.RotateDegrees(-90);
            Etichetta.Disegna(c, etichette[j], sfondo);
            c.Restore();
        }

        using var nero = new SKPaint { Color = SKColors.Black, IsAntialias = true };
        float r = (float)(6.0 / 25.4 * Dpi / 2);
        float sx = (float)(MargineX / 2), dx = (float)(PaginaW - MargineX / 2);
        foreach (var (cx, fy) in new[] { (sx, 0.2017), (sx, 0.9717), (dx, 0.0901), (dx, 0.6361) })
            c.DrawCircle(cx, (float)(PaginaH * fy), r, nero);

        using var linea = new SKPaint
        {
            Color = SKColors.Black, Style = SKPaintStyle.Stroke, StrokeWidth = 3, StrokeCap = SKStrokeCap.Butt,
        };
        float vx = (float)(PaginaW * 0.99764), vy = (float)(PaginaH * 0.99356);
        c.DrawLine(vx - (float)(PaginaW * 0.005805), vy, vx, vy, linea);
        c.DrawLine(vx, vy - (float)(PaginaH * 0.009442), vx, vy, linea);
    }
}
