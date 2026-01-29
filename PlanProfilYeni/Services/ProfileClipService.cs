// Services/ProfileClipService.cs
using System;
using System.Collections.Generic;
using PlanProfilYeni.Domain;

namespace PlanProfilYeni.Services
{
    /// <summary>
    /// Class7 km kırığı motoru.
    /// SADECE Arazi ve Boru profilleri içindir.
    /// Hidrolik bu servisi KULLANMAZ.
    /// </summary>
    public sealed class ProfileClipService
    {
        public List<ProfilePolylineBandSegment> SplitByBands(
            double[] kmElevationPairs,
            ProfileBandSet bands)
        {
            if (kmElevationPairs == null) throw new ArgumentNullException(nameof(kmElevationPairs));
            if (bands == null) throw new ArgumentNullException(nameof(bands));

            var breaks = GetEffectiveBreaks(bands.BreakKms);
            var points = ReadPoints(kmElevationPairs);

            foreach (double b in breaks)
                InsertBreakpoint(points, b);

            return BuildBandSegments(points, breaks, bands.BandCount);
        }

        // --- AŞAĞISI AYNI, SADECE YORUMLAR DÜZELTİLDİ ---

        private static double[] GetEffectiveBreaks(double[] breakKms)
        {
            var list = new List<double>();
            foreach (var x in breakKms)
                if (x > 0 && x < 999999.0)
                    list.Add(x);
            return list.ToArray();
        }

        private struct Pt { public double X, Y; public Pt(double x, double y) { X = x; Y = y; } }

        private static List<Pt> ReadPoints(double[] pairs)
        {
            var pts = new List<Pt>();
            for (int i = 0; i < pairs.Length; i += 2)
                pts.Add(new Pt(pairs[i], pairs[i + 1]));
            return pts;
        }

        private static void InsertBreakpoint(List<Pt> points, double breakKm)
        {
            for (int i = 1; i < points.Count; i++)
            {
                var p1 = points[i - 1];
                var p2 = points[i];

                if (p1.X < breakKm && breakKm < p2.X)
                {
                    double y = p1.Y + (breakKm - p1.X) * (p2.Y - p1.Y) / (p2.X - p1.X);
                    points.Insert(i, new Pt(breakKm, y));
                    return;
                }
            }
        }

        private static List<ProfilePolylineBandSegment> BuildBandSegments(
            List<Pt> points, double[] breaks, int bandCount)
        {
            var result = new List<ProfilePolylineBandSegment>();
            int band = 0, bi = 0;
            double nextBreak = breaks.Length > 0 ? breaks[0] : double.MaxValue;

            var buf = new List<double>();
            buf.Add(points[0].X); buf.Add(points[0].Y);

            for (int i = 1; i < points.Count; i++)
            {
                var p = points[i];

                while (p.X > nextBreak)
                {
                    Add(result, band, buf);
                    buf = new List<double> { nextBreak, buf[buf.Count - 1] };
                    band = Math.Min(band + 1, bandCount - 1);
                    bi++;
                    nextBreak = bi < breaks.Length ? breaks[bi] : double.MaxValue;
                }

                buf.Add(p.X); buf.Add(p.Y);
            }

            Add(result, band, buf);
            return result;
        }

        private static void Add(List<ProfilePolylineBandSegment> res, int band, List<double> buf)
        {
            if (buf.Count >= 4)
                res.Add(new ProfilePolylineBandSegment(band, buf.ToArray()));
        }
    }
}
