using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using PlanProfilYeni.Cad;
using PlanProfilYeni.Services;
using System;
using System.Collections.Generic;
using PlanProfilYeni.Domain;

namespace PlanProfilYeni.Application
{
    public sealed class PrintHydraulicTableToCadUseCase
    {
        private readonly HydraulicExcelService _excelService;

        public PrintHydraulicTableToCadUseCase(HydraulicExcelService excelService)
        {
            if (excelService == null) throw new ArgumentNullException(nameof(excelService));
            _excelService = excelService;
        }

        public void Execute(AcadDocument doc, HydraulicProcessOptions options, CadHydraulicTableOptions cadOptions)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (cadOptions == null) throw new ArgumentNullException(nameof(cadOptions));

            if (options.InputFiles == null || options.InputFiles.Length == 0)
                throw new InvalidOperationException("Hiç Excel dosyası seçilmedi.");

            List<HydraulicRow> rows = _excelService.Read(options.InputFiles);
            if (rows == null || rows.Count == 0)
                throw new InvalidOperationException("Excel dosyalarından veri okunamadı.");

            var printer = new CadHydraulicTablePrinter(doc, cadOptions);
            printer.Print(rows);
        }
    }
}
