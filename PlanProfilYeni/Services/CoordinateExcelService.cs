using Excel = Microsoft.Office.Interop.Excel;
using PlanProfilYeni.Domain;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace PlanProfilYeni.Services
{
    public class CoordinateExcelService
    {
        public List<CoordinateRow> Read(string[] files)
        {
            var list = new List<CoordinateRow>();

            foreach (var file in files)
                list.AddRange(ReadSingle(file));

            return list;
        }

        private List<CoordinateRow> ReadSingle(string path)
        {
            var rows = new List<CoordinateRow>();

            Excel.Application excel = new();
            excel.Visible = false;
            excel.DisplayAlerts = false;

            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                wb = excel.Workbooks.Open(path);
                ws = wb.Worksheets["Boru Plan Koordinatları"];

                int row = 3;
                double km = 0;

                while (!string.IsNullOrWhiteSpace(ws.Cells[row, "H"].Value?.ToString()))
                {
                    string hat = Path.GetFileNameWithoutExtension(path);
                    string nokta;

                    if (row == 3)
                        nokta = "Hat Başı";
                    else if (string.IsNullOrWhiteSpace(ws.Cells[row + 1, "H"].Value?.ToString()))
                        nokta = "Hat Sonu";
                    else
                        nokta = "S" + (row - 3);

                    double x = Convert.ToDouble(ws.Cells[row, "H"].Value);
                    double y = Convert.ToDouble(ws.Cells[row, "I"].Value);

                    string donus = "-";
                    string sapma = "-";

                    if (row > 3 && !string.IsNullOrWhiteSpace(ws.Cells[row + 1, "H"].Value?.ToString()))
                    {
                        double x1 = Convert.ToDouble(ws.Cells[row - 1, "H"].Value);
                        double y1 = Convert.ToDouble(ws.Cells[row - 1, "I"].Value);
                        double x2 = x;
                        double y2 = y;
                        double x3 = Convert.ToDouble(ws.Cells[row + 1, "H"].Value);
                        double y3 = Convert.ToDouble(ws.Cells[row + 1, "I"].Value);

                        var (d, s) = GeometryService.CalculateAngles(x1, y1, x2, y2, x3, y3);

                        donus = d.ToString("0.00");
                        sapma = s.ToString("0.00");
                    }

                    rows.Add(new CoordinateRow
                    {
                        HatAdi = hat,
                        NoktaAdi = nokta,
                        X = x.ToString("### ### ###.000"),
                        Y = y.ToString("### ### ###.000"),
                        Acisal1 = sapma,
                        Acisal2 = donus,
                        Km = km.ToString("0+000.00")
                    });

                    km += Convert.ToDouble(ws.Cells[row + 1, "E"].Value);
                    row++;
                }
            }
            finally
            {
                wb?.Close(false);
                excel.Quit();

                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(excel);
            }

            return rows;
        }
    }
}
