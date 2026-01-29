// Services/HydraulicVerticalBreakService.cs
using System.Collections.Generic;

namespace PlanProfilYeni.Services
{
    /// <summary>
    /// Class7 hidrolik band içi düşey kırık motoru.
    /// Arazi/Boru için ASLA kullanılmaz.
    /// </summary>
    public sealed class HydraulicVerticalBreakService
    {
        private const double Step = 4.0;

        public List<double[]> BreakVertically(
            double[] kmPressurePairs,
            double topLimit)
        {
            var result = new List<double[]>();
            var buf = new List<double>();

            double currentTop = topLimit;

            for (int i = 0; i < kmPressurePairs.Length; i += 2)
            {
                double km = kmPressurePairs[i];
                double val = kmPressurePairs[i + 1];

                while (val < currentTop - Step)
                {
                    result.Add(buf.ToArray());
                    buf.Clear();
                    currentTop -= Step;
                }

                buf.Add(km);
                buf.Add(val);
            }

            if (buf.Count >= 4)
                result.Add(buf.ToArray());

            return result;
        }
    }
}
