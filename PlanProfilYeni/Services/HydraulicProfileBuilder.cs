// Services/HydraulicProfileBuilder.cs
using System;
using System.Collections.Generic;
using PlanProfilYeni.Domain;

namespace PlanProfilYeni.Services
{
    /// <summary>
    /// Hidrolik profile ait (km,value) serisini:
    /// - band km kırıklarına (ProfileBandSet.BreakKms) göre böler,
    /// - band içinde topLimit (TopRefMeters) düşürerek 4m adımlı mikro kırık uygular,
    /// - çizilebilir segmentler + kırık marker'ları üretir.
    /// </summary>
    public sealed class HydraulicProfileBuilder
    {
        private struct Pt { public double X, Y; public Pt(double x, double y) { X = x; Y = y; } }

        public sealed class BuildResult
        {
            public List<HydraulicDrawableSegment> Segments { get; private set; }
            public List<HydraulicVerticalBreakMarker> VerticalBreakMarkers { get; private set; }

            public BuildResult(List<HydraulicDrawableSegment> segs, List<HydraulicVerticalBreakMarker> markers)
            {
                Segments = segs ?? new List<HydraulicDrawableSegment>();
                VerticalBreakMarkers = markers ?? new List<HydraulicVerticalBreakMarker>();
            }
        }

        /// <summary>
        /// </summary>
        /// <param name="kmValuePairs">[km0,val0, km1,val1,...] (val: hidrolik değer)</param>
        /// <param name="pipeKmElevationPairs">[km,kot] boru profili (D/E) — gerekirse ValueMode=AddToPipeElevationMeters için kullanılır</param>
        public BuildResult Build(
            double[] kmValuePairs,
            double[] pipeKmElevationPairs,
            ProfileBandSet bands,
            HydraulicBuildOptions options)
        {
            if (kmValuePairs == null) throw new ArgumentNullException(nameof(kmValuePairs));
            if (pipeKmElevationPairs == null) throw new ArgumentNullException(nameof(pipeKmElevationPairs));
            if (bands == null) throw new ArgumentNullException(nameof(bands));
            if (options == null) throw new ArgumentNullException(nameof(options));

            if (kmValuePairs.Length < 4 || (kmValuePairs.Length % 2) != 0)
                throw new ArgumentException("kmValuePairs must be [km,val] pairs.", nameof(kmValuePairs));
            if (pipeKmElevationPairs.Length < 4 || (pipeKmElevationPairs.Length % 2) != 0)
                throw new ArgumentException("pipeKmElevationPairs must be [km,kot] pairs.", nameof(pipeKmElevationPairs));

            options.Validate();

            var pts = ReadPoints(kmValuePairs, pipeKmElevationPairs, options);

            // Band geçişleri: BreakKms listesi (sentinel dahil olabilir)
            var bandBreaks = GetEffectiveBreaks(bands.BreakKms);

            var segs = new List<HydraulicDrawableSegment>(64);
            var markers = new List<HydraulicVerticalBreakMarker>(64);

            int bandIndex = 0;
            int breakPtr = 0;
            double bandEndKm = (bandBreaks.Length > 0) ? bandBreaks[0] : double.MaxValue;

            double topRef = bands.TopLevels[bandIndex]; // initial topLimit
            double step = options.VerticalBreakStepMeters;

            var buf = new List<double>(128);

            // Start point
            buf.Add(pts[0].X);
            buf.Add(pts[0].Y);

            for (int i = 1; i < pts.Count; i++)
            {
                Pt p = pts[i];

                // 1) Band km kırığı (ikincil/koşullu ama gerekli)
                // Nokta band sonunu geçtiyse, önce band kesişim noktasını üretip bandı ilerletmeliyiz.
                while (p.X > bandEndKm)
                {
                    // bandEndKm'de bir kesişim noktası üret (interpolate)
                    Pt prev = GetPrevPointForInterpolation(pts, i);
                    Pt cut = new Pt(bandEndKm, InterpolateY(prev, p, bandEndKm));

                    // Bu noktayı mevcut segmente ekle
                    buf.Add(cut.X);
                    buf.Add(cut.Y);

                    // Segmenti finalize et
                    AddSegmentIfValid(segs, bandIndex, topRef, buf, bands.TopLevels[bandIndex]);

                    // Yeni banda geç: aynı bandEndKm noktasından yeni buffer başlat
                    buf = new List<double>(128);
                    buf.Add(cut.X);
                    buf.Add(cut.Y);

                    // band advance
                    bandIndex = Math.Min(bandIndex + 1, bands.BandCount - 1);
                    breakPtr++;
                    bandEndKm = (breakPtr < bandBreaks.Length) ? bandBreaks[breakPtr] : double.MaxValue;

                    // Class7: yeni banda geçince topRef reset (bandın top level'ı)
                    topRef = bands.TopLevels[bandIndex];
                }

                // 2) Band içi düşey mikro kırık (hidrolik için öncelikli davranış)
                // Class7: while (value < topLimit - 4.0) { topLimit -= 4.0; kırık çizgisi bas; segmenti kes; }
                while (p.Y < topRef - step)
                {
                    double from = topRef;
                    double to = topRef - step;

                    // marker (KotDeğişimÇizgisi) km konumu: Class7 pratikte "o anki x" civarı
                    markers.Add(new HydraulicVerticalBreakMarker(bandIndex, p.X, from, to));

                    // Segmenti kes (mevcut buffer'ı finalize)
                    AddSegmentIfValid(segs, bandIndex, topRef, buf, bands.TopLevels[bandIndex]);

                    // topRef güncelle
                    topRef = to;

                    // yeni buffer: aynı noktadan devam (Class7’de kırık noktası her iki parçada da bulunur)
                    buf = new List<double>(128);
                    buf.Add(p.X);
                    buf.Add(p.Y);
                }

                // Normal nokta ekle
                buf.Add(p.X);
                buf.Add(p.Y);
            }

            AddSegmentIfValid(segs, bandIndex, topRef, buf, bands.TopLevels[bandIndex]);
            return new BuildResult(segs, markers);
        }

        private static Pt GetPrevPointForInterpolation(List<Pt> pts, int currentIndex)
        {
            // currentIndex>=1 garantili
            return pts[currentIndex - 1];
        }

        private static List<Pt> ReadPoints(double[] kmValuePairs, double[] pipeKmElevationPairs, HydraulicBuildOptions opt)
        {
            var pts = new List<Pt>(kmValuePairs.Length / 2);

            for (int i = 0; i < kmValuePairs.Length; i += 2)
            {
                double km = kmValuePairs[i];
                double val = kmValuePairs[i + 1];

                double elevation;
                if (opt.ValueMode == HydraulicValueMode.AbsoluteElevationMeters)
                {
                    elevation = val;
                }
                else
                {
                    // pipeElev + (val * factor)
                    double pipeElev = InterpolateYAtX(pipeKmElevationPairs, km);
                    elevation = pipeElev + val * opt.ValueToMetersFactor;
                }

                pts.Add(new Pt(km, elevation));
            }

            return pts;
        }

        private static double[] GetEffectiveBreaks(double[] breakKms)
        {
            var list = new List<double>(breakKms.Length);
            for (int i = 0; i < breakKms.Length; i++)
            {
                double x = breakKms[i];
                if (x <= 0) continue;
                if (x >= 999999.0) continue;
                list.Add(x);
            }
            return list.ToArray();
        }

        private static double InterpolateY(Pt p1, Pt p2, double x)
        {
            double dx = p2.X - p1.X;
            if (Math.Abs(dx) < 1e-12) return p1.Y;
            double t = (x - p1.X) / dx;
            return p1.Y + t * (p2.Y - p1.Y);
        }

        private static double InterpolateYAtX(double[] xyPairs, double x)
        {
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

            if (x <= xyPairs[0]) return xyPairs[1];
            return xyPairs[xyPairs.Length - 1];
        }

        private static void AddSegmentIfValid(List<HydraulicDrawableSegment> segs, int bandIndex, double topRef, List<double> buf, double bandTopLevel)
        {
            if (buf == null || buf.Count < 4) return;

            // Degenerate temizliği: start==end ise çizme
            double x0 = buf[0];
            double y0 = buf[1];
            double x1 = buf[buf.Count - 2];
            double y1 = buf[buf.Count - 1];
            if (Math.Abs(x1 - x0) < 1e-12 && Math.Abs(y1 - y0) < 1e-12)
                return;

            var seg = new HydraulicDrawableSegment(bandIndex, topRef, buf.ToArray());
            seg.BandTopLevelMeters = bandTopLevel;  // ✅ BandTopLevel'i set et
            segs.Add(seg);
        }
    }
}
