using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace PlanProfilYeni.Services
{
    /// <summary>
    /// Arazi & Boru Profili Excel okuma servisi.
    /// GeometryService ile KARIŞMAZ.
    /// </summary>
    public sealed class ProfileExcelService
    {
        /// <summary>
        /// Arazi profili: B/C sütunları, satır 4+
        /// </summary>

        public double[] ReadGroundProfile(Excel.Worksheet ws)
        {
            var list = new List<double>();

            for (int r = 4; ws.Cells[r, "B"].Value != null; r++)
            {
                double km = 0;
                double kot = 0;

                if (TryDouble(ws.Cells[r, "B"].Value, out km) &&
                    TryDouble(ws.Cells[r, "C"].Value, out kot))
                {
                    list.Add(km);
                    list.Add(kot);
                }
            }

            return list.ToArray();
        }

        /// <summary>
        /// Boru profili: D/E sütunları, satır 4+
        /// Eğer D sütunu satır 4'te diktik değilse, B sütunundaki son satırı referans al (arazi ile sync).
        /// </summary>
        public double[] ReadPipeProfile(Excel.Worksheet ws)
        {
            var list = new List<double>();

            // Arazi profilinin son satırını bul (B sütunu)
            int maxRow = 4;
            for (int r = 4; ws.Cells[r, "B"].Value != null; r++)
            {
                maxRow = r;
            }

            // D/E'yi satır 4 ile maxRow arasında oku
            for (int r = 4; r <= maxRow; r++)
            {
                double km = 0;
                double kot = 0;

                object dVal = ws.Cells[r, "D"].Value;
                object eVal = ws.Cells[r, "E"].Value;

                // D/E boş satırı geç
                if (dVal == null && eVal == null)
                    continue;

                if (TryDouble(dVal, out km) &&
                    TryDouble(eVal, out kot))
                {
                    list.Add(km);
                    list.Add(kot);
                }
            }

            return list.ToArray();
        }

        private static bool TryDouble(object v, out double d)
        {
            d = 0;
            if (v == null) return false;
            return double.TryParse(v.ToString(), out d);
        }
    }
}
