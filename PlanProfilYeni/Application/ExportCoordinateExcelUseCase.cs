using PlanProfilYeni.Services;
using System;

namespace PlanProfilYeni.Application
{
    public class ExportCoordinateExcelUseCase
    {
        private readonly CoordinateExcelService _reader = new();
        private readonly CoordinateReportWriter _writer = new();

        public void Execute(string[] files)
        {
            var rows = _reader.Read(files);

            if (rows.Count == 0)
                throw new Exception("Veri bulunamadı");

            _writer.Export(rows);
        }
    }
}
