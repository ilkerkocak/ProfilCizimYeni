using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common; // Referanslarda olmalı
using PlanProfilYeni.Cad;
using PlanProfilYeni.Services;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices; // Marshal için gerekli
using PlanProfilYeni.Domain;

namespace PlanProfilYeni.Application
{
    public sealed class PrintHydraulicTableToCadUseCase
    {
        private readonly HydraulicExcelService _excelService;

        // Constructor
        public PrintHydraulicTableToCadUseCase(HydraulicExcelService excelService)
        {
            if (excelService == null) throw new ArgumentNullException(nameof(excelService));
            _excelService = excelService;
        }

        // Metodu Form1'den gelen yapıya uygun hale getirdik:
        // Artık (string[] dosyalar, CadHydraulicTableOptions ayarlar) alıyor.
        public void Execute(string[] excelFiles, CadHydraulicTableOptions cadOptions)
        {
            // 1. Parametre Kontrolü
            if (excelFiles == null || excelFiles.Length == 0)
                throw new InvalidOperationException("Hiç Excel dosyası seçilmedi.");

            if (cadOptions == null)
                throw new ArgumentNullException(nameof(cadOptions));

            // 2. Excel Verilerini Oku
            List<HydraulicRow> rows = _excelService.Read(excelFiles);

            if (rows == null || rows.Count == 0)
                throw new InvalidOperationException("Excel dosyalarından veri okunamadı.");

            // 3. AutoCAD Bağlantısını Kur (Form1 yerine burada yapıyoruz)
            AcadApplication acadApp;
            try
            {
                // Açık olan AutoCAD'i yakalar
                acadApp = (AcadApplication)Marshal.GetActiveObject("AutoCAD.Application");
            }
            catch
            {
                throw new InvalidOperationException("AutoCAD uygulaması açık değil. Lütfen AutoCAD'i açıp tekrar deneyin.");
            }

            AcadDocument doc = acadApp.ActiveDocument;

            // 4. Çizim İşlemini Başlat
            var printer = new CadHydraulicTablePrinter(doc, cadOptions);
            printer.Print(rows);
        }
    }
}