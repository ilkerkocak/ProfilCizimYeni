using System;

namespace PlanProfilYeni.Services
{
    public static class GeometryService
    {
        public static (double donusGrad, double sapmaGrad) CalculateAngles(
            double x1, double y1,
            double x2, double y2,
            double x3, double y3)
        {
            double v1x = x2 - x1;
            double v1y = y2 - y1;

            double v2x = x3 - x2;
            double v2y = y3 - y2;

            double a1 = Math.Atan2(v1y, v1x);
            double a2 = Math.Atan2(v2y, v2x);

            double delta = a1 - a2;
            double grad = delta * 400.0 / (2.0 * Math.PI);

            if (grad > 200) grad -= 400;
            else if (grad < -200) grad += 400;

            double sapma = 200.0 - grad;

            if (sapma < 0) sapma += 400;
            else if (sapma > 400) sapma -= 400;

            return (Math.Round(grad, 2), Math.Round(sapma, 2));
        }
    }
}
