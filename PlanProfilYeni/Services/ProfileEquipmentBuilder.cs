// Services/ProfileEquipmentBuilder.cs
using System;
using System.Collections.Generic;
using PlanProfilYeni.Domain;

namespace PlanProfilYeni.Services
{
    /// <summary>
    /// Class7 profil ekipman listesini birleştirir.
    /// - Hidrant: (km=N, outlet=BH)
    /// - HatAyrimi: (km=N, label=BD prefix)
    /// - BKV: (km=N, staticWaterLevel opsiyonel)
    /// - Vantuz/Tahliye: boru extrema + manuel listeler (B12.. / C12..)
    /// 
    /// Not: Excel okuma burada değil; Excel servisleri ham listeleri üretir.
    /// </summary>
    public sealed class ProfileEquipmentBuilder
    {
        /// <summary>
        /// Class7’de vantuz listesi, tahliye listesiyle 1.0 km tolerans içinde çakışıyorsa vantuzdan siliniyor.
        /// </summary>
        public double VantuzTahliyeConflictKmTolerance { get; set; } = 1.0;

        public List<ProfileEquipmentItem> Build(
            IList<Tuple<double, int>> hydrants,                 // (km, outletCount)  BH
            IList<Tuple<double, string>> hatAyrimlari,           // (km, label)       BD " Ayr."
            IList<Tuple<double, double?>> bkvs,                  // (km, ssk)         BD "BKV" + I ilişkisi
            IList<double> vantuzFromPipe,                        // boru maxima km
            IList<double> tahliyeFromPipe,                       // boru minima km
            IList<double> manualVantuzKms,                       // Basınç Profili B12..
            IList<double> manualTahliyeKms)                      // Basınç Profili C12..
        {
            var items = new List<ProfileEquipmentItem>(128);

            // 1) Hidrant
            if (hydrants != null)
            {
                for (int i = 0; i < hydrants.Count; i++)
                {
                    var t = hydrants[i];
                    if (t != null)
                        items.Add(ProfileEquipmentItem.Hydrant(t.Item1, t.Item2));
                }
            }

            // 2) Hat ayrımı
            if (hatAyrimlari != null)
            {
                for (int i = 0; i < hatAyrimlari.Count; i++)
                {
                    var t = hatAyrimlari[i];
                    if (t != null && !string.IsNullOrWhiteSpace(t.Item2))
                        items.Add(ProfileEquipmentItem.HatAyrimi(t.Item1, t.Item2));
                }
            }

            // 3) BKV
            if (bkvs != null)
            {
                for (int i = 0; i < bkvs.Count; i++)
                {
                    var t = bkvs[i];
                    if (t != null)
                        items.Add(ProfileEquipmentItem.Bkv(t.Item1, t.Item2));
                }
            }

            // 4) Vantuz/Tahliye: boru extrema + manuel listeler
            var vantuz = new List<double>(64);
            var tahliye = new List<double>(64);

            AddAll(vantuz, vantuzFromPipe);
            AddAll(tahliye, tahliyeFromPipe);

            // Class7: B12.. aralıktaki değerleri uygun aralığa düşüyorsa vantuza ekliyor.
            // Biz burada doğrudan ekliyoruz (aralık filtresi boru extrema servisinde veya Excel servisinde de yapılabilir).
            AddAll(vantuz, manualVantuzKms);

            // Class7: C12.. okuyor; sonra vantuz listesinde 1.0 km yakın olanları siliyor.
            AddAll(tahliye, manualTahliyeKms);
            RemoveConflictingVantuz(vantuz, tahliye, VantuzTahliyeConflictKmTolerance);

            // items'a ekle
            for (int i = 0; i < vantuz.Count; i++)
                items.Add(ProfileEquipmentItem.Vantuz(vantuz[i]));

            for (int i = 0; i < tahliye.Count; i++)
                items.Add(ProfileEquipmentItem.Tahliye(tahliye[i]));

            return items;
        }

        private static void AddAll(List<double> target, IList<double> src)
        {
            if (src == null) return;
            for (int i = 0; i < src.Count; i++)
            {
                double km = src[i];
                if (km > 0) target.Add(km);
            }
        }

        private static void RemoveConflictingVantuz(List<double> vantuz, List<double> tahliye, double tol)
        {
            if (vantuz.Count == 0 || tahliye.Count == 0) return;

            // Class7: vantuz listesinde geriye doğru gezip, tahliye listesinde 1.0 km yakınlık varsa vantuzu siliyor.
            for (int i = vantuz.Count - 1; i >= 0; i--)
            {
                double v = vantuz[i];
                for (int j = 0; j < tahliye.Count; j++)
                {
                    if (Math.Abs(v - tahliye[j]) < tol)
                    {
                        vantuz.RemoveAt(i);
                        break;
                    }
                }
            }
        }
    }
}
