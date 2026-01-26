using PlanProfilYeni.Services;
using System;
using System.Collections.Generic;
using PlanProfilYeni.Domain;

namespace PlanProfilYeni.Application
{
    public class ImportExcelTablesUseCase
    {
        private readonly HydraulicExcelService _excelReader;
        private readonly HydraulicReportWriter _reportWriter;

        public ImportExcelTablesUseCase()
        {
            _excelReader = new HydraulicExcelService();
            _reportWriter = new HydraulicReportWriter();
        }

        public void Execute(string[] inputFiles, string outputPath)
        {
            if (inputFiles == null || inputFiles.Length == 0)
                throw new InvalidOperationException("Excel dosyası seçilmedi.");

            if (string.IsNullOrWhiteSpace(outputPath))
                throw new InvalidOperationException("Çıktı yolu belirtilmedi.");

            var rows = _excelReader.Read(inputFiles);   // ✅ doğru alan

            if (rows.Count == 0)
                throw new InvalidOperationException("Veri okunamadı.");

            _reportWriter.Export(rows, outputPath, false);
        }
    }
}
