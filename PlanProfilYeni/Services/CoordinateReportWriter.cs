using Excel = Microsoft.Office.Interop.Excel;
using PlanProfilYeni.Domain;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace PlanProfilYeni.Services
{
    public class CoordinateReportWriter
    {
        public void Export(List<CoordinateRow> rows)
        {
            var excel = new Excel.Application { Visible = true };

            var wb = excel.Workbooks.Add();
            var ws = wb.ActiveSheet;
            ws.Name = "Koordinat Tablosu";

            string[] headers = { "Hat", "Nokta", "X", "Y", "Açı1", "Açı2", "Km" };

            for (int i = 0; i < headers.Length; i++)
                ws.Cells[1, i + 1] = headers[i];

            int r = 2;
            foreach (var row in rows)
            {
                ws.Cells[r, 1] = row.HatAdi;
                ws.Cells[r, 2] = row.NoktaAdi;
                ws.Cells[r, 3] = row.X;
                ws.Cells[r, 4] = row.Y;
                ws.Cells[r, 5] = row.Acisal1;
                ws.Cells[r, 6] = row.Acisal2;
                ws.Cells[r, 7] = row.Km;
                r++;
            }
        }
    }
}
