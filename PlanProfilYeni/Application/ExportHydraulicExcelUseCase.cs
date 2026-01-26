using PlanProfilYeni.Services;
using System;
using System.Collections.Generic;
using PlanProfilYeni.Domain;

namespace PlanProfilYeni.Application
{
    public sealed class ExportHydraulicExcelUseCase
    {
        private readonly HydraulicExcelService _excelService;
        private readonly HydraulicReportWriter _reportWriter;

        public ExportHydraulicExcelUseCase(HydraulicExcelService excelService, HydraulicReportWriter reportWriter)
        {
            if (excelService == null) throw new ArgumentNullException(nameof(excelService));
            if (reportWriter == null) throw new ArgumentNullException(nameof(reportWriter));

            _excelService = excelService;
            _reportWriter = reportWriter;
        }

        public void Execute(HydraulicProcessOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (options.InputFiles == null || options.InputFiles.Length == 0)
                throw new InvalidOperationException("Hiç Excel dosyası seçilmedi.");

            // 1) Oku
            List<HydraulicRow> rows = _excelService.Read(options.InputFiles);
            if (rows == null || rows.Count == 0)
                throw new InvalidOperationException("Excel dosyalarından veri okunamadı.");

            // 2) Yaz (outputPath boş olabilir, openAfterExport true ise kullanıcı kaydeder)
            _reportWriter.Export(rows, options.OutputExcelPath, options.OpenExcelAfterExport);
        }
    }
}
