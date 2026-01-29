// Services/ProfileBandService.cs
using System;
using System.Collections.Generic;
using PlanProfilYeni.Domain;

namespace PlanProfilYeni.Services
{
    /// <summary>
    /// Class7 smethod_7 içerisindeki band üretimi + break km üretiminin birebir karşılığı.
    /// Girdi: groundXY ve pipeXY = [x0,y0,x1,y1,...] (km, kot) dizileri.
    /// </summary>
    public sealed class ProfileBandService
    {
        // Class7: for (double num5 = 0.1; num5 <= num4; num5 += 0.1)
        private const double ScanStepKm = 0.1;

        // Class7: if (num2 - num3 > 13)
        private const int BandHeight = 13;

        // Class7: gparam_5 en sona 999999.0
        private const double BreakKmSentinel = 999999.0;

        public ProfileBandSet BuildBands(double[] groundXY, double[] pipeXY, string lineName)
        {
            if (groundXY == null) throw new ArgumentNullException(nameof(groundXY));
            if (pipeXY == null) throw new ArgumentNullException(nameof(pipeXY));
            if (groundXY.Length < 4 || groundXY.Length % 2 != 0)
                throw new ArgumentException("groundXY must be [x,y] pairs, length >= 4.", nameof(groundXY));
            if (pipeXY.Length < 4 || pipeXY.Length % 2 != 0)
                throw new ArgumentException("pipeXY must be [x,y] pairs, length >= 4.", nameof(pipeXY));

            // Class7: int num2 = (int)Math.Round(Math.Ceiling(double_0[1]));
            //         int num3 = (int)Math.Round(Math.Floor(double_1[1]));
            int num2 = (int)Math.Round(Math.Ceiling(groundXY[1])); // top candidate
            int num3 = (int)Math.Round(Math.Floor(pipeXY[1]));     // base candidate

            // Class7: double num4 = double_1[UBound(double_1) - 1];
            // Gerçek max km değerini hesapla (hem ground hem pipe için)
            double maxGroundKm = groundXY[groundXY.Length - 2];
            double maxPipeKm = pipeXY[pipeXY.Length - 2];
            double endKm = Math.Max(maxGroundKm, maxPipeKm);

            // 1) Band sınırlarını üret (gparam_3/gparam_4)
            var topLevels = new List<int>(16);
            var baseLevels = new List<int>(16);

            for (double km = ScanStepKm; km <= endKm; km += ScanStepKm)
            {
                bool flag = false;   // ground raised top
                bool flag2 = false;  // pipe lowered base

                double groundElev = InterpolateYAtX(groundXY, km);
                double pipeElev = InterpolateYAtX(pipeXY, km);

                // Class7: if (num10 <= num7) else MessageBox -> çizme
                if (pipeElev > groundElev)
                {
                    // Class7 "Boru Arazinin üstünde" => çizim iptal.
                    throw new InvalidOperationException(
                        "Dikkat, Boru Arazinin üstünde. Hat:" + lineName + " KM= " + km.ToString("0.0") + " Bu hat çizilmeyecek!");
                }

                if (groundElev > num2)
                {
                    num2 = (int)Math.Round(Math.Ceiling(groundElev));
                    flag = true;
                }

                if (pipeElev < num3)
                {
                    num3 = (int)Math.Round(Math.Floor(pipeElev));
                    flag2 = true;
                }

                if (num2 - num3 > BandHeight)
                {
                    // Class7: if (flag2) { add (num2, num2-13) } then reset
                    if (flag2)
                    {
                        topLevels.Add(num2);
                        baseLevels.Add(num2 - BandHeight);

                        num2 = (int)Math.Round(Math.Ceiling(groundElev));
                        num3 = (int)Math.Round(Math.Floor(pipeElev));
                    }

                    // Class7: if (flag) { add (num3+13, num3) } then reset
                    if (flag)
                    {
                        topLevels.Add(num3 + BandHeight);
                        baseLevels.Add(num3);

                        num2 = (int)Math.Round(Math.Ceiling(groundElev));
                        num3 = (int)Math.Round(Math.Floor(pipeElev));
                    }
                }
            }

            // Class7: smethod_1(ref gparam_3, num2); smethod_1(ref gparam_4, num2 - 13);
            // NOTE: Class7 her durumda son banda (top=num2, base=num2-13) ekliyor.
            topLevels.Add(num2);
            baseLevels.Add(num2 - BandHeight);

            // 2) Break km (gparam_5) üret (band sınır kesişimleri)
            double[] breakKms = BuildBreakKms(groundXY, pipeXY, topLevels, baseLevels, endKm);

            return new ProfileBandSet(topLevels.ToArray(), baseLevels.ToArray(), breakKms);
        }

        /// <summary>
        /// Class7: gparam_5 üretimi.
        /// - Boru segmenti baseKot altına düşerse: baseKot kesişim km
        /// - Arazi noktası topKot üstüne çıkarsa: topKot kesişim km
        /// - En sona gerçek max km eklenir.
        /// </summary>
        private static double[] BuildBreakKms(double[] groundXY, double[] pipeXY, List<int> topLevels, List<int> baseLevels, double endKm)
        {
            int k = 0; // band index (Class7 'k')

            var breaks = new List<double>(Math.Max(8, topLevels.Count + 2));

            // Class7: for (l=2; l<=UBound(pipe); l+=2)
            for (int l = 2; l <= pipeXY.Length - 2; l += 2)
            {
                double x1 = pipeXY[l - 2];
                double x2 = pipeXY[l];
                double y1 = pipeXY[l - 1];
                double y2 = pipeXY[l + 1];

                // Class7: if (pipeY2 < baseLevels[k]) { for(; pipeY2 < baseLevels[k]; k++) { ... add xBreak } }
                if (y2 < baseLevels[k])
                {
                    while (y2 < baseLevels[k])
                    {
                        // num16 = (y1 - y2) / (x2 - x1)
                        double num16 = (y1 - y2) / (x2 - x1);

                        // xBreak = x1 + 1/num16 * (y1 - baseKot)
                        double xBreak = x1 + (1.0 / num16) * (y1 - baseLevels[k]);
                        breaks.Add(xBreak);

                        k++;
                        if (k >= baseLevels.Count)
                        {
                            // Defensive: Class7 sentinel ile zaten çizimde koruyor; burada sınır taşmasını engelle.
                            k = baseLevels.Count - 1;
                            break;
                        }
                    }
                    continue;
                }

                // Class7: arazi noktalarını tarıyor: m=2.., eğer (groundX[m] > x1 && groundX[m] <= x2) && groundY(m) > topLevels[k] ...
                for (int m = 2; m <= groundXY.Length - 2; m += 2)
                {
                    double gx = groundXY[m];
                    double gy = groundXY[m + 1];

                    if ((gx > x1) && (gx <= x2) && (gy > topLevels[k]))
                    {
                        while (gy > topLevels[k])
                        {
                            // num18 = (groundY(m) - groundY(m-2? actually m-1)) / (groundX(m) - groundX(m-2))
                            // Class7 uses: num18 = (double_0[m + 1] - double_0[m - 1]) / (double_0[m] - double_0[m - 2])
                            double num18 = (groundXY[m + 1] - groundXY[m - 1]) / (groundXY[m] - groundXY[m - 2]);

                            // xBreak = groundX(m-2) + 1/num18 * (topKot - groundY(m-1))
                            double xBreak2 = groundXY[m - 2] + (1.0 / num18) * (topLevels[k] - groundXY[m - 1]);
                            breaks.Add(xBreak2);

                            k++;
                            if (k >= topLevels.Count)
                            {
                                k = topLevels.Count - 1;
                                break;
                            }
                        }
                    }
                }
            }

            // Break km sayısı, band sayısına eşit olmalı
            // Eğer eksik break km varsa, geri kalan band'ler için endKm kullan
            while (breaks.Count < topLevels.Count)
            {
                breaks.Add(endKm);
            }

            // Eğer fazla break km varsa (çok sayıda kesişim), sadece band sayısı kadar al
            // ve en sonuncuyu endKm yap
            if (breaks.Count > topLevels.Count)
            {
                breaks.RemoveRange(topLevels.Count, breaks.Count - topLevels.Count);
            }

            // Son break km her zaman gerçek maksimum km olmalı
            if (breaks.Count > 0)
            {
                breaks[breaks.Count - 1] = endKm;
            }
            else
            {
                // Hiç break km üretilmemişse (tek band), endKm ekle
                breaks.Add(endKm);
            }

            return breaks.ToArray();
        }

        /// <summary>
        /// Polyline üzerinde (x,y) çiftlerinden x=km için y enterpolasyonu.
        /// Class7 taramada: segment bul, slope=(y2-y1)/(x2-x1), y=y1+(x-x1)*slope.
        /// </summary>
        private static double InterpolateYAtX(double[] xyPairs, double x)
        {
            // Class7 loop: i=2..ubound step2, if (x[i-2] < x && x[i] > x) ...
            for (int i = 2; i <= xyPairs.Length - 2; i += 2)
            {
                double x1 = xyPairs[i - 2];
                double x2 = xyPairs[i];
                if (x1 < x && x2 > x)
                {
                    double y1 = xyPairs[i - 1];
                    double y2 = xyPairs[i + 1];
                    double slope = (y2 - y1) / (x2 - x1);
                    return y1 + (x - x1) * slope;
                }
            }

            // Boundary handling: Class7 taraması genelde (0.1..endKm) ve veriler kapsayıcı.
            // Yine de emniyetli davranalım: en yakın uç noktayı döndür.
            if (x <= xyPairs[0]) return xyPairs[1];
            return xyPairs[xyPairs.Length - 1];
        }
    }
}
