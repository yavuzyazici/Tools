using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

class ListDiff
{
    static void Main(string[] args)
    {
        // ============================================================
        // DOSYA YOLLARI
        // ============================================================

        string downloadsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "Downloads"
        );

        string registrar = "Marcaria";

        string pathA = Path.Combine(downloadsPath, "registrar.txt");
        string pathB = Path.Combine(downloadsPath, "dna.txt");

        if (!File.Exists(pathA)) { Console.WriteLine($"HATA: Dosya bulunamadı → {pathA}"); Console.ReadKey(); return; }
        if (!File.Exists(pathB)) { Console.WriteLine($"HATA: Dosya bulunamadı → {pathB}"); Console.ReadKey(); return; }

        Console.WriteLine("Dosyalar okunuyor...");

        var listA = File.ReadAllLines(pathA)
            .Select(x => x.Trim())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();

        var listB = File.ReadAllLines(pathB)
            .Select(x => x.Trim())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .ToList();

        // ============================================================
        // FARK HESAPLA (büyük/küçük harf duyarsız, boşluk temizle)
        // ============================================================

        var setA = new HashSet<string>(listA.Select(x => x.Trim().ToLowerInvariant()));
        var setB = new HashSet<string>(listB.Select(x => x.Trim().ToLowerInvariant()));

        var onlyInA = listA
            .Select(x => x.Trim())
            .Where(x => !setB.Contains(x.ToLowerInvariant()))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(x => x)
            .ToList();

        var onlyInB = listB
            .Select(x => x.Trim())
            .Where(x => !setA.Contains(x.ToLowerInvariant()))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(x => x)
            .ToList();

        Console.WriteLine($"Sadece {registrar}'da: {onlyInA.Count} öğe");
        Console.WriteLine($"Sadece DNA'da: {onlyInB.Count} öğe");

        // ============================================================
        // EXCEL'E KAYDET
        // ============================================================

        string outputPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "Downloads",
            $"{string.Concat(registrar.Split(' ').Select(w => char.ToUpper(w[0]) + w.Substring(1)))}ListeFarklari.xlsx"
        );

        using var wb = new XLWorkbook();

        // --- Sayfa 1: Sadece A'da ---
        var wsA = wb.AddWorksheet($"Sadece {registrar}'da");
        wsA.Cell(1, 1).Value = $"Sadece {registrar} Listesinde Olanlar";
        wsA.Cell(1, 1).Style.Font.Bold = true;
        wsA.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#FF6B6B");
        wsA.Cell(1, 1).Style.Font.FontColor = XLColor.White;

        for (int i = 0; i < onlyInA.Count; i++)
            wsA.Cell(i + 2, 1).Value = onlyInA[i];

        wsA.Column(1).Width = 45;

        // --- Sayfa 2: Sadece B'de ---
        var wsB = wb.AddWorksheet("Sadece DNA'da");
        wsB.Cell(1, 1).Value = "Sadece DNA Listesinde Olanlar";
        wsB.Cell(1, 1).Style.Font.Bold = true;
        wsB.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#51CF66");
        wsB.Cell(1, 1).Style.Font.FontColor = XLColor.White;

        for (int i = 0; i < onlyInB.Count; i++)
            wsB.Cell(i + 2, 1).Value = onlyInB[i];

        wsB.Column(1).Width = 45;

        // --- Sayfa 3: Özet ---
        var wsSummary = wb.AddWorksheet("Özet");
        wsSummary.Cell(1, 1).Value = "Kategori";
        wsSummary.Cell(1, 2).Value = "Adet";
        wsSummary.Cell(1, 1).Style.Font.Bold = true;
        wsSummary.Cell(1, 2).Style.Font.Bold = true;

        wsSummary.Cell(2, 1).Value = $"{registrar} Listesi Toplam";
        wsSummary.Cell(2, 2).Value = listA.Count;

        wsSummary.Cell(3, 1).Value = "DNA Listesi Toplam";
        wsSummary.Cell(3, 2).Value = listB.Count;

        wsSummary.Cell(4, 1).Value = $"Sadece {registrar}'da";
        wsSummary.Cell(4, 2).Value = onlyInA.Count;
        wsSummary.Cell(4, 1).Style.Font.FontColor = XLColor.FromHtml("#C92A2A");

        wsSummary.Cell(5, 1).Value = "Sadece DNA'da";
        wsSummary.Cell(5, 2).Value = onlyInB.Count;
        wsSummary.Cell(5, 1).Style.Font.FontColor = XLColor.FromHtml("#2F9E44");

        wsSummary.Cell(6, 1).Value = "Her İkisinde De (Ortak)";
        wsSummary.Cell(6, 2).Value = setA.Intersect(setB).Count();

        wsSummary.Columns().AdjustToContents();

        wb.SaveAs(outputPath);

        Console.WriteLine($"\nExcel kaydedildi: {outputPath}");
        Console.ReadKey();
    }
}