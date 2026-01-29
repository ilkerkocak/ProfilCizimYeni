using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using PlanProfilYeni.Domain;

namespace PlanProfilYeni.Cad
{
    /// <summary>
    /// Hidrolik profil çizer:
    /// - Arazi/boru gridinin 2m üstünde ayrı bir 4m'lik grid kutusu
    /// - Hidrolik değerleri (örn: 1070-1074) bu 4m'lik banda normalize edilir
    /// - Kendi grid çizgileri ve koordinat sistemi
    /// </summary>
    public sealed class CadProfileHydraulicDrawer
    {
        private readonly AcadDocument _doc;
        private readonly ProfileBandSet _bands;

        // Hidrolik grid arazi gridinden 2m yukarıda başlar
        private const double VERTICAL_OFFSET_METERS = 2.0;
        // Hidrolik grid yüksekliği 4 metre
        private const double HYDRAULIC_GRID_HEIGHT = 4.0;

        public string GridThinLayer { get; set; } = "Hidrolik-Grid-ince";
        public string GridThickLayer { get; set; } = "Hidrolik-Grid-kalin";
        public string CurveLayer { get; set; } = "Hidrolik-Eğri";
        public string BreakLayer { get; set; } = "KotDeğişimÇizgisi";

        public CadProfileHydraulicDrawer(AcadDocument doc, ProfileBandSet bands)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _bands = bands ?? throw new ArgumentNullException(nameof(bands));
        }

        public void Draw(
            IEnumerable<HydraulicDrawableSegment> segments,
            IEnumerable<HydraulicVerticalBreakMarker> breakMarkers,
            CadProfileTransformer transformer,
            CadProfileTransformOptions transformOptions)
        {
            if (segments == null) throw new ArgumentNullException(nameof(segments));
            if (breakMarkers == null) throw new ArgumentNullException(nameof(breakMarkers));
            if (transformer == null) throw new ArgumentNullException(nameof(transformer));
            if (transformOptions == null) throw new ArgumentNullException(nameof(transformOptions));

            AcadModelSpace ms = _doc.ModelSpace;

            try
            {
                // Layerları oluştur
                EnsureLayer("Hidrolik-Grid-ince", 254);   // Açık gri
                EnsureLayer("Hidrolik-Grid-kalin", 150);  // Koyu gri
                EnsureLayer("Hidrolik-Eğri", 6);          // Magenta
                EnsureLayer("KotDeğişimÇizgisi", 253);    // Gri

                // Hidrolik değerlerinin min/max kotunu bul
                double minElev, maxElev;
                FindHydraulicElevationRange(segments, out minElev, out maxElev);

                // Her band için hidrolik grid ve eğrileri çiz
                for (int b = 0; b < _bands.BandCount; b++)
                {
                    double startKm = (b == 0) ? 0.0 : _bands.BreakKms[b - 1];
                    double endKm = _bands.BreakKms[b];

                    // Hidrolik grid kutusu çiz
                    DrawHydraulicGrid(ms, b, startKm, endKm, minElev, maxElev, transformer, transformOptions);

                    // Bu banda ait hidrolik eğrileri çiz
                    DrawHydraulicCurvesForBand(ms, segments, b, minElev, maxElev, transformer, transformOptions);
                }

                // Mikro kırık işaretleri
                DrawBreakMarkers(ms, breakMarkers, minElev, maxElev, transformer, transformOptions);
            }
            finally
            {
                ReleaseCom(ms);
            }
        }

        private void FindHydraulicElevationRange(
            IEnumerable<HydraulicDrawableSegment> segments,
            out double minElev,
            out double maxElev)
        {
            minElev = double.MaxValue;
            maxElev = double.MinValue;

            foreach (var seg in segments)
            {
                var pairs = seg.KmElevationPairs;
                for (int i = 1; i < pairs.Length; i += 2)
                {
                    double elev = pairs[i];
                    if (elev < minElev) minElev = elev;
                    if (elev > maxElev) maxElev = elev;
                }
            }

            // Eğer veri yoksa varsayılan
            if (minElev == double.MaxValue)
            {
                minElev = 1070;
                maxElev = 1074;
            }
        }

        private void DrawHydraulicGrid(
            AcadModelSpace ms,
            int bandIndex,
            double startKm,
            double endKm,
            double minElev,
            double maxElev,
            CadProfileTransformer transformer,
            CadProfileTransformOptions transformOptions)
        {
            // CAD koordinatları: Arazi gridinin üstünde
            double xLeft = transformer.ToCadX(startKm, bandIndex);
            double xRight = transformer.ToCadX(endKm, bandIndex);

            // Arazi gridinin üst sınırı
            double araziTopY = transformer.ToCadY(_bands.TopLevels[bandIndex], bandIndex);

            // Hidrolik grid 2m yukarıda başlar
            double offsetCad = VERTICAL_OFFSET_METERS * transformOptions.CadUnitsPerMeter;
            double hydraulicBottomY = araziTopY + offsetCad;
            double hydraulicTopY = hydraulicBottomY + (HYDRAULIC_GRID_HEIGHT * transformOptions.CadUnitsPerMeter);

            // Border (kalın çerçeve)
            AddLine(ms, xLeft, hydraulicBottomY, xRight, hydraulicBottomY, GridThickLayer);
            AddLine(ms, xRight, hydraulicBottomY, xRight, hydraulicTopY, GridThickLayer);
            AddLine(ms, xRight, hydraulicTopY, xLeft, hydraulicTopY, GridThickLayer);
            AddLine(ms, xLeft, hydraulicTopY, xLeft, hydraulicBottomY, GridThickLayer);

            // Yatay ince gridler (20cm)
            DrawHorizontalLines(ms, xLeft, xRight, hydraulicBottomY, hydraulicTopY, 0.2, transformOptions, GridThinLayer);

            // Yatay kalın gridler (1m)
            DrawHorizontalLines(ms, xLeft, xRight, hydraulicBottomY, hydraulicTopY, 1.0, transformOptions, GridThickLayer);

            // Dikey gridler (100m)
            DrawVerticalLines(ms, bandIndex, startKm, endKm, hydraulicBottomY, hydraulicTopY, transformer, GridThickLayer);
        }

        private void DrawHorizontalLines(
            AcadModelSpace ms,
            double xLeft,
            double xRight,
            double yBottom,
            double yTop,
            double stepMeters,
            CadProfileTransformOptions transformOptions,
            string layer)
        {
            double stepCad = stepMeters * transformOptions.CadUnitsPerMeter;
            double firstY = yBottom + stepCad;

            int count = (int)Math.Floor((yTop - yBottom) / stepCad);
            if (count <= 0) return;

            AcadLine ln = null;
            try
            {
                ln = (AcadLine)ms.AddLine(
                    new double[] { xLeft, firstY, 0.0 },
                    new double[] { xRight, firstY, 0.0 });
                ln.Layer = layer;

                if (count > 1)
                    ln.ArrayRectangular(count, 1, 1, stepCad, 1.0, 1.0);
            }
            finally
            {
                ReleaseCom(ln);
            }
        }

        private void DrawVerticalLines(
            AcadModelSpace ms,
            int bandIndex,
            double startKm,
            double endKm,
            double yBottom,
            double yTop,
            CadProfileTransformer transformer,
            string layer)
        {
            double stepKm = 100.0;
            double firstKm = NextMultiple(startKm, stepKm);
            if (firstKm >= endKm) return;

            double xFirst = transformer.ToCadX(firstKm, bandIndex);
            double xNext = transformer.ToCadX(firstKm + stepKm, bandIndex);
            double spacing = Math.Abs(xNext - xFirst);

            if (spacing <= 0.000001) return;

            double xEnd = transformer.ToCadX(endKm, bandIndex);
            int count = (int)Math.Floor(Math.Abs(xEnd - xFirst) / spacing) + 1;
            if (count <= 0) return;

            AcadLine ln = null;
            try
            {
                ln = (AcadLine)ms.AddLine(
                    new double[] { xFirst, yBottom, 0.0 },
                    new double[] { xFirst, yTop, 0.0 });
                ln.Layer = layer;

                if (count > 1)
                    ln.ArrayRectangular(1, count, 1, 1.0, spacing, 1.0);
            }
            finally
            {
                ReleaseCom(ln);
            }
        }

        private void DrawHydraulicCurvesForBand(
            AcadModelSpace ms,
            IEnumerable<HydraulicDrawableSegment> segments,
            int bandIndex,
            double minElev,
            double maxElev,
            CadProfileTransformer transformer,
            CadProfileTransformOptions transformOptions)
        {
            // Arazi gridinin üst sınırı
            double araziTopY = transformer.ToCadY(_bands.TopLevels[bandIndex], bandIndex);

            // Hidrolik grid 2m yukarıda başlar
            double offsetCad = VERTICAL_OFFSET_METERS * transformOptions.CadUnitsPerMeter;
            double hydraulicBottomY = araziTopY + offsetCad;
            double hydraulicGridHeightCad = HYDRAULIC_GRID_HEIGHT * transformOptions.CadUnitsPerMeter;

            // Kot aralığı
            double elevRange = maxElev - minElev;
            if (elevRange <= 0) elevRange = 4.0;

            foreach (var seg in segments)
            {
                if (seg.BandIndex != bandIndex) continue;

                var pairs = seg.KmElevationPairs;

                for (int i = 2; i < pairs.Length; i += 2)
                {
                    double km1 = pairs[i - 2];
                    double e1 = pairs[i - 1];
                    double km2 = pairs[i];
                    double e2 = pairs[i + 1];

                    double x1 = transformer.ToCadX(km1, bandIndex);
                    double x2 = transformer.ToCadX(km2, bandIndex);

                    // Normalize: minElev → 0, maxElev → HYDRAULIC_GRID_HEIGHT
                    double normalizedE1 = ((e1 - minElev) / elevRange) * HYDRAULIC_GRID_HEIGHT;
                    double normalizedE2 = ((e2 - minElev) / elevRange) * HYDRAULIC_GRID_HEIGHT;

                    // CAD Y koordinatı
                    double y1 = hydraulicBottomY + (normalizedE1 * transformOptions.CadUnitsPerMeter);
                    double y2 = hydraulicBottomY + (normalizedE2 * transformOptions.CadUnitsPerMeter);

                    AcadLine ln = null;
                    try
                    {
                        ln = (AcadLine)ms.AddLine(
                            new double[] { x1, y1, 0 },
                            new double[] { x2, y2, 0 });
                        ln.Layer = CurveLayer;
                    }
                    finally
                    {
                        ReleaseCom(ln);
                    }
                }
            }
        }

        private void DrawBreakMarkers(
            AcadModelSpace ms,
            IEnumerable<HydraulicVerticalBreakMarker> breakMarkers,
            double minElev,
            double maxElev,
            CadProfileTransformer transformer,
            CadProfileTransformOptions transformOptions)
        {
            if (breakMarkers == null) return;

            double elevRange = maxElev - minElev;
            if (elevRange <= 0) elevRange = 4.0;

            foreach (var mk in breakMarkers)
            {
                int bandIndex = mk.BandIndex;
                double x = transformer.ToCadX(mk.Km, bandIndex);

                // Arazi gridinin üst sınırı
                double araziTopY = transformer.ToCadY(_bands.TopLevels[bandIndex], bandIndex);

                // Hidrolik grid 2m yukarıda başlar
                double offsetCad = VERTICAL_OFFSET_METERS * transformOptions.CadUnitsPerMeter;
                double hydraulicBottomY = araziTopY + offsetCad;

                // Top ve Bottom kotları normalize et
                double normalizedTop = ((mk.FromTopRefMeters - minElev) / elevRange) * HYDRAULIC_GRID_HEIGHT;
                double normalizedBot = ((mk.ToTopRefMeters - minElev) / elevRange) * HYDRAULIC_GRID_HEIGHT;

                double yTop = hydraulicBottomY + (normalizedTop * transformOptions.CadUnitsPerMeter);
                double yBot = hydraulicBottomY + (normalizedBot * transformOptions.CadUnitsPerMeter);

                AcadLine ln = null;
                try
                {
                    ln = (AcadLine)ms.AddLine(
                        new double[] { x, yTop, 0 },
                        new double[] { x, yBot, 0 });
                    ln.Layer = BreakLayer;
                }
                finally
                {
                    ReleaseCom(ln);
                }
            }
        }

        private void AddLine(AcadModelSpace ms, double x1, double y1, double x2, double y2, string layer)
        {
            AcadLine ln = null;
            try
            {
                ln = (AcadLine)ms.AddLine(
                    new double[] { x1, y1, 0 },
                    new double[] { x2, y2, 0 });
                ln.Layer = layer;
            }
            finally
            {
                ReleaseCom(ln);
            }
        }

        private double NextMultiple(double val, double mult)
        {
            return Math.Ceiling(val / mult) * mult;
        }

        private void EnsureLayer(string layerName, int colorIndex)
        {
            try
            {
                var layer = _doc.Layers.Item(layerName);
                layer.color = (AcColor)colorIndex;
            }
            catch
            {
                var layer = _doc.Layers.Add(layerName);
                layer.color = (AcColor)colorIndex;
            }
        }

        private static void ReleaseCom(object o)
        {
            try
            {
                if (o != null && Marshal.IsComObject(o))
                    Marshal.ReleaseComObject(o);
            }
            catch
            {
            }
        }
    }
}
