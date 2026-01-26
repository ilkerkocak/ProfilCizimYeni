using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using PlanProfilYeni.Domain;

namespace PlanProfilYeni.Services
{
    public sealed class HydraulicExcelService
    {
        private const int DataStartRow = 8;
        private const string SheetName = "Basınç Profili";

        public List<HydraulicRow> Read(string[] filePaths)
        {
            if (filePaths == null) throw new ArgumentNullException(nameof(filePaths));

            var all = new List<HydraulicRow>();
            for (int i = 0; i < filePaths.Length; i++)
            {
                var path = filePaths[i];
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                    continue;

                all.AddRange(ReadSingle(path));
            }
            return all;
        }

        private List<HydraulicRow> ReadSingle(string filePath)
        {
            var rows = new List<HydraulicRow>();

            Excel.Application excel = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excel = new Excel.Application { Visible = false, DisplayAlerts = false };
                workbook = excel.Workbooks.Open(filePath);
                worksheet = workbook.Worksheets[SheetName];

                int row = DataStartRow;

                // N sütunu boş olana kadar oku
                while (!IsEmpty(worksheet.Cells[row, "N"] != null ? worksheet.Cells[row, "N"].Value : null))
                {
                    var model = ReadRow(worksheet, row, Path.GetFileNameWithoutExtension(filePath));
                    if (model != null) rows.Add(model);
                    row++;
                }

                return rows;
            }
            finally
            {
                ReleaseComObject(worksheet);

                if (workbook != null)
                {
                    workbook.Close(false);
                    ReleaseComObject(workbook);
                }

                if (excel != null)
                {
                    excel.Quit();
                    ReleaseComObject(excel);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private HydraulicRow ReadRow(Excel.Worksheet ws, int row, string hatAdi)
        {
            // Boru tipi
            string boru = ToText(ws.Cells[row, "T"] != null ? ws.Cells[row, "T"].Value : null);

            // İç / dış çap (PE’de V iç, W dış; diğerlerinde ters)
            double? v = TryDouble(ws.Cells[row, "V"] != null ? ws.Cells[row, "V"].Value : null);
            double? w = TryDouble(ws.Cells[row, "W"] != null ? ws.Cells[row, "W"].Value : null);

            double? icCap = (boru == "PE") ? v : w;
            double? disCap = (boru == "PE") ? w : v;

            // Dinamik basınç BJ
            double? dinamik = TryDouble(ws.Cells[row, "BJ"] != null ? ws.Cells[row, "BJ"].Value : null);
            bool? regulator = dinamik.HasValue ? (bool?)(dinamik.Value >= 5.0) : null;
            bool? limiter = dinamik.HasValue ? (bool?)(dinamik.Value >= 40.0) : null;

            // Hidrant detay alanı (BF boşsa eski koddaki gibi detaylar “-” sayılır)
            double? hizmetAlan = TryDouble(ws.Cells[row, "BF"] != null ? ws.Cells[row, "BF"].Value : null);

            var model = new HydraulicRow();
            model.HatAdi = hatAdi;
            model.KesitNo = row - (DataStartRow - 1);
            model.Kilometre = FormatKm(
                ws.Cells[row, "L"] != null ? ws.Cells[row, "L"].Value : null,
                ws.Cells[row, "N"] != null ? ws.Cells[row, "N"].Value : null
            );

            model.BrutAlan = TryDouble(ws.Cells[row, "O"] != null ? ws.Cells[row, "O"].Value : null);
            model.SulamaModulu = TryDouble(ws.Cells[row, "P"] != null ? ws.Cells[row, "P"].Value : null);
            model.HatDebi = TryDouble(ws.Cells[row, "Q"] != null ? ws.Cells[row, "Q"].Value : null);
            model.Hiz = TryDouble(ws.Cells[row, "R"] != null ? ws.Cells[row, "R"].Value : null);
            model.HidrolikEğim = TryDouble(ws.Cells[row, "S"] != null ? ws.Cells[row, "S"].Value : null);

            model.BoruCinsi = boru;
            model.IcCap = icCap;
            model.DisCap = disCap;

            model.StatikBasinc = Round2(ws.Cells[row, "BB"] != null ? ws.Cells[row, "BB"].Value : null);
            model.DayanmaBasinci = TryDouble(ws.Cells[row, "BC"] != null ? ws.Cells[row, "BC"].Value : null);
            model.HidrantNo = ToText(ws.Cells[row, "BD"] != null ? ws.Cells[row, "BD"].Value : null);

            model.HizmetAlan = hizmetAlan;
            model.Debi = TryDouble(ws.Cells[row, "BG"] != null ? ws.Cells[row, "BG"].Value : null);
            model.CikisSayisi = TryInt(ws.Cells[row, "BH"] != null ? ws.Cells[row, "BH"].Value : null);
            model.Tipi = ToText(ws.Cells[row, "BI"] != null ? ws.Cells[row, "BI"].Value : null);

            model.DinamikBasinc = dinamik;
            model.BasincRegulatoru = regulator;
            model.DebiLimitoru = limiter;

            return model;
        }

        private static string FormatKm(object start, object end)
        {
            double s, e;
            if (!TryParseAnyDouble(start, out s)) return "";
            if (!TryParseAnyDouble(end, out e)) return "";
            return string.Format(CultureInfo.InvariantCulture, "{0:0+000.00} - {1:0+000.00}", s, e);
        }

        private static double? TryDouble(object value)
        {
            if (value == null) return null;

            if (value is double) return (double)value;
            if (value is float) return Convert.ToDouble(value);
            if (value is int) return Convert.ToDouble(value);
            if (value is long) return Convert.ToDouble(value);
            if (value is decimal) return Convert.ToDouble(value);

            var s = value.ToString();
            if (string.IsNullOrWhiteSpace(s)) return null;

            // TR/EN karışık dosyalara tolerans: önce TR, sonra Invariant
            double d;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out d)) return d;
            if (double.TryParse(s.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out d)) return d;

            return null;
        }

        private static bool TryParseAnyDouble(object value, out double result)
        {
            result = 0.0;
            var v = TryDouble(value);
            if (!v.HasValue) return false;
            result = v.Value;
            return true;
        }

        private static int? TryInt(object value)
        {
            if (value == null) return null;

            if (value is int) return (int)value;
            if (value is double) return (int)(double)value;

            var s = value.ToString();
            if (string.IsNullOrWhiteSpace(s)) return null;

            int i;
            if (int.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out i)) return i;

            //  "12,0" gibi gelirse:
            double d;
            if (double.TryParse(s.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                return (int)d;

            return null;
        }

        private static double? Round2(object value)
        {
            var d = TryDouble(value);
            if (!d.HasValue) return null;
            return Math.Round(d.Value, 2);
        }

        private static string ToText(object value)
        {
            if (value == null) return null;
            var s = value.ToString();
            return string.IsNullOrWhiteSpace(s) ? null : s.Trim();
        }

        private static bool IsEmpty(object value)
        {
            if (value == null) return true;
            var s = value.ToString();
            return string.IsNullOrWhiteSpace(s);
        }

        private static void ReleaseComObject(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
                Marshal.ReleaseComObject(obj);
        }
    }
}
