// Services/PipeExtremaDetector.cs
using System;
using System.Collections.Generic;

namespace PlanProfilYeni.Services
{
    /// <summary>
    /// Boru profili (km,kot) üzerinden yerel maksimum/minimumları bulur.
    /// Class7: basit üçlü karşılaştırma, filtre yok.
    /// </summary>
    public sealed class PipeExtremaDetector
    {
        public void Detect(
            double[] pipeKmElevationPairs,
            out List<double> vantuzKms,
            out List<double> tahliyeKms)
        {
            vantuzKms = new List<double>();
            tahliyeKms = new List<double>();

            if (pipeKmElevationPairs == null || pipeKmElevationPairs.Length < 6)
                return;

            int n = pipeKmElevationPairs.Length / 2;

            for (int i = 1; i < n - 1; i++)
            {
                double kmPrev = pipeKmElevationPairs[(i - 1) * 2];
                double yPrev = pipeKmElevationPairs[(i - 1) * 2 + 1];

                double kmCur = pipeKmElevationPairs[i * 2];
                double yCur = pipeKmElevationPairs[i * 2 + 1];

                double kmNext = pipeKmElevationPairs[(i + 1) * 2];
                double yNext = pipeKmElevationPairs[(i + 1) * 2 + 1];

                // Vantuz: yerel maksimum
                if (yPrev < yCur && yCur > yNext)
                    vantuzKms.Add(kmCur);

                // Tahliye: yerel minimum
                if (yPrev > yCur && yCur < yNext)
                    tahliyeKms.Add(kmCur);
            }

            // Class7: uç noktalar opsiyonel; burada eklemiyoruz (gerekirse parametrelenir)
        }
    }
}
