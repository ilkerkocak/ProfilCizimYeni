using Autodesk.AutoCAD.Interop;
using PlanProfilYeni.Cad;
using PlanProfilYeni.Services;
using System;
using System.Runtime.InteropServices;

namespace PlanProfilYeni.Application
{
    public class PrintCoordinateCadUseCase
    {
        private readonly CoordinateExcelService _reader = new CoordinateExcelService();

        public void Execute(string[] files)
        {
            if (files == null || files.Length == 0)
                throw new InvalidOperationException("Hiç Excel dosyası seçilmedi.");

            // AutoCAD bağlantısını kur
            AcadApplication acadApp;
            try
            {
                acadApp = (AcadApplication)Marshal.GetActiveObject("AutoCAD.Application");
            }
            catch
            {
                throw new InvalidOperationException("AutoCAD uygulaması açık değil. Lütfen AutoCAD'i açıp tekrar deneyin.");
            }

            AcadDocument doc = acadApp.ActiveDocument;

            if (doc == null)
                throw new Exception("AutoCAD bağlı değil.");

            var rows = _reader.Read(files);

            var printer = new CadCoordinateTablePrinter(doc);
            printer.Print(rows);
        }
    }
}