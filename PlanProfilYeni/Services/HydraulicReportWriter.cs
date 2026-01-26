using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using PlanProfilYeni.Domain;

namespace PlanProfilYeni.Services
{
    public sealed class HydraulicReportWriter
    {
        public void Export(List<HydraulicRow> rows, string outputPath, bool openAfterExport)
        {
            if (rows == null) throw new ArgumentNullException(nameof(rows));

            Excel.Application excel = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                excel = new Excel.Application { Visible = false, DisplayAlerts = false };
                wb = excel.Workbooks.Add();
                ws = (Excel.Worksheet)wb.ActiveSheet;
                ws.Name = "Hidrolik Tablo";

                WriteHeaders(ws);
                WriteRows(ws, rows);

                // Otomatik genişlik
                ws.Columns.AutoFit();

                if (!string.IsNullOrWhiteSpace(outputPath))
                {
                    wb.SaveAs(outputPath);
                }

                if (openAfterExport)
                {
                    // kullanıcı kaydetsin diye ekranda aç
                    excel.Visible = true;
                    // COM nesnelerini burada serbest bırakmıyoruz; Excel açık kalacak.
                    // Kullanıcı kapatınca süreç biter.
                    return;
                }

                wb.Close(false);
            }
            finally
            {
                // openAfterExport ise excel'i kapatmıyoruz
                // (kullanıcı Excel penceresini görmek istiyor).
                // Bu durumda COM cleanup yapılmaz; Excel kapanınca sistem serbest bırakır.
                // openAfterExport false ise cleanup yaparız.
            }

            // cleanup (sadece Excel kapalıyken güvenli)
            ReleaseCom(ws);
            ReleaseCom(wb);
            if (excel != null)
            {
                excel.Quit();
                ReleaseCom(excel);
            }
        }

        private static void WriteHeaders(Excel.Worksheet ws)
        {
            string[] headers =
            {
                "Hat", "Kesit", "Kilometre", "Brüt Alan", "Modül", "Hat Debi",
                "Hız", "Eğim", "Boru", "İç Çap", "Dış Çap", "Statik",
                "Dayanma", "Hidrant", "Hizmet Alan", "Debi", "Çıkış", "Tip", "Dinamik",
                "Basınç Regülatörü", "Debi Limitörü"
            };

            for (int i = 0; i < headers.Length; i++)
                ws.Cells[1, i + 1] = headers[i];

            // başlık formatı
            ws.Rows[1].Font.Bold = true;
        }

        private static void WriteRows(Excel.Worksheet ws, List<HydraulicRow> rows)
        {
            var ci = CultureInfo.InvariantCulture;

            int row = 2;
            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];

                ws.Cells[row, 1] = r.HatAdi ?? "-";
                ws.Cells[row, 2] = r.KesitNo;
                ws.Cells[row, 3] = r.Kilometre ?? "-";

                ws.Cells[row, 4] = r.BrutAlan;
                ws.Cells[row, 5] = r.SulamaModulu;
                ws.Cells[row, 6] = r.HatDebi;
                ws.Cells[row, 7] = r.Hiz;
                ws.Cells[row, 8] = r.HidrolikEğim;

                ws.Cells[row, 9] = r.BoruCinsi ?? "-";
                ws.Cells[row, 10] = r.IcCap;
                ws.Cells[row, 11] = r.DisCap;

                ws.Cells[row, 12] = r.StatikBasinc;
                ws.Cells[row, 13] = r.DayanmaBasinci;

                ws.Cells[row, 14] = r.HidrantNo ?? "-";
                ws.Cells[row, 15] = r.HizmetAlan;
                ws.Cells[row, 16] = r.Debi;
                ws.Cells[row, 17] = r.CikisSayisi;
                ws.Cells[row, 18] = r.Tipi ?? "-";
                ws.Cells[row, 19] = r.DinamikBasinc;

                // Reg / Lim (Var/Yok/-)
                ws.Cells[row, 20] = r.BasincRegulatoru.HasValue ? (r.BasincRegulatoru.Value ? "VAR" : "YOK") : "-";
                ws.Cells[row, 21] = r.DebiLimitoru.HasValue ? (r.DebiLimitoru.Value ? "VAR" : "YOK") : "-";

                row++;
            }
        }

        private static void ReleaseCom(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
                Marshal.ReleaseComObject(obj);
        }
    }
}
