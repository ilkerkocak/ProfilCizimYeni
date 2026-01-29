// Services/PressureProfileExcelService.cs
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace PlanProfilYeni.Services
{
    public sealed partial class PressureProfileExcelService
    {
        /// <summary>
        /// Hidrantlar: Km=N, Çıkış sayısı=BH (satır 8+)
        /// </summary>
        public List<Tuple<double, int>> ReadHydrants(Excel.Worksheet ws)
        {
            var list = new List<Tuple<double, int>>();
            for (int r = 8; ws.Cells[r, "N"].Value != null; r++)
            {
                object kmObj = ws.Cells[r, "N"].Value;
                object bhObj = ws.Cells[r, "BH"].Value;

                if (kmObj == null || bhObj == null) continue;

                double km;
                int outlet;
                if (TryDouble(kmObj, out km) && TryInt(bhObj, out outlet))
                {
                    list.Add(Tuple.Create(km, outlet));
                }
            }
            return list;
        }

        /// <summary>
        /// Hat ayrımları: Km=N, metin=BD (BD içinde " Ayr." geçenler)
        /// </summary>
        public List<Tuple<double, string>> ReadHatAyrimlari(Excel.Worksheet ws)
        {
            var list = new List<Tuple<double, string>>();
            for (int r = 8; ws.Cells[r, "N"].Value != null; r++)
            {
                object kmObj = ws.Cells[r, "N"].Value;
                object bdObj = ws.Cells[r, "BD"].Value;

                if (kmObj == null || bdObj == null) continue;

                double km;
                string text = bdObj.ToString();
                if (TryDouble(kmObj, out km) && text.Contains(" Ayr."))
                {
                    // Class7: "xxx Ayr." öncesi etiket
                    string label = text.Substring(0, text.IndexOf(" Ayr.", StringComparison.Ordinal)).Trim();
                    list.Add(Tuple.Create(km, label));
                }
            }
            return list;
        }

        /// <summary>
        /// BKV: Km=N, tanım=BD starts "BKV"
        /// SSK ilişkisi: Cells[r-4,"I"] (Class7)
        /// </summary>
        public List<Tuple<double, double?>> ReadBkvs(Excel.Worksheet ws)
        {
            var list = new List<Tuple<double, double?>>();
            for (int r = 8; ws.Cells[r, "N"].Value != null; r++)
            {
                object kmObj = ws.Cells[r, "N"].Value;
                object bdObj = ws.Cells[r, "BD"].Value;

                if (kmObj == null || bdObj == null) continue;

                double km;
                string text = bdObj.ToString().Trim();
                if (TryDouble(kmObj, out km) && text.StartsWith("BKV", StringComparison.OrdinalIgnoreCase))
                {
                    double? ssk = null;
                    object sskObj = ws.Cells[r - 4, "I"].Value; // Class7 ilişkisi
                    double sskVal;
                    if (TryDouble(sskObj, out sskVal))
                        ssk = sskVal;

                    list.Add(Tuple.Create(km, ssk));
                }
            }
            return list;
        }

        /// <summary>
        /// Manuel Vantuz Km'leri: B12.. boş gelene kadar
        /// </summary>
        public List<double> ReadManualVantuzKms(Excel.Worksheet ws)
        {
            var list = new List<double>();
            for (int r = 12; ws.Cells[r, "B"].Value != null; r++)
            {
                double km;
                if (TryDouble(ws.Cells[r, "B"].Value, out km))
                    list.Add(km);
            }
            return list;
        }

        /// <summary>
        /// Manuel Tahliye Km'leri: C12.. boş gelene kadar
        /// </summary>
        public List<double> ReadManualTahliyeKms(Excel.Worksheet ws)
        {
            var list = new List<double>();
            for (int r = 12; ws.Cells[r, "C"].Value != null; r++)
            {
                double km;
                if (TryDouble(ws.Cells[r, "C"].Value, out km))
                    list.Add(km);
            }
            return list;
        }
        private static bool TryDouble(object v, out double d)
        {
            d = 0;
            if (v == null) return false;
            return double.TryParse(v.ToString(), out d);
        }

        private static bool TryInt(object v, out int i)
        {
            i = 0;
            if (v == null) return false;
            return int.TryParse(v.ToString(), out i);
        }
        public double[] ReadHydraulicSeries(Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            var list = new List<double>();

            // Başlangıç noktası
            double km0 = 0, kot0 = 0;
            if (TryDouble(ws.Cells[4, "B"].Value, out km0) &&
                TryDouble(ws.Cells[4, "C"].Value, out kot0))
            {
                list.Add(km0);
                list.Add(kot0);
            }

            // Devamı: N8 / X8 ...
            for (int r = 8; ws.Cells[r, "N"].Value != null; r++)
            {
                double km = 0, kot = 0;
                if (TryDouble(ws.Cells[r, "N"].Value, out km) &&
                    TryDouble(ws.Cells[r, "X"].Value, out kot))
                {
                    list.Add(km);
                    list.Add(kot);
                }
            }

            return list.ToArray();
        }
    }
}
