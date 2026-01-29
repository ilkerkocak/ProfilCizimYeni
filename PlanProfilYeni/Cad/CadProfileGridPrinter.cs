using System;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;

namespace PlanProfilYeni.Cad
{
    /// <summary>
    /// Class7 grid yaklaşımı:
    /// - Çok sayıda AddLine yerine ArrayRectangular kullan.
    /// - Major/minor grid için Grid-kalın / Grid-ince layer.
    ///
    /// Backward-compatible API:
    /// GridLayer, VerticalGridStepMeters, HorizontalGridStepMeters, MaxVerticalLines, MaxHorizontalLines
    /// </summary>
    public sealed class CadProfileGridPrinter
    {
        private readonly AcadDocument _doc;

        // --- Class7 layerları (önerilen) ---
        public string GridThinLayer { get; set; } = "Grid-ince";
        public string GridThickLayer { get; set; } = "Grid-kalın";
        public string VerticalGridLayer { get; set; } = "KmÇizgisi";  // Ara dikey gridler için

        // --- Grid adımları (metre) ---
        public double VerticalStepMeters { get; set; } = 100.0;      // 100 m
        public double HorizontalThinStepMeters { get; set; } = 1.0;   // 1 m (ince)
        public double HorizontalThickStepMeters { get; set; } = 5.0;  // 5 m (kalın)

        // --- Emniyet limitleri ---
        public int MaxVerticalLinesPerBand { get; set; } = 400;
        public int MaxHorizontalThinLinesPerBand { get; set; } = 800;
        public int MaxHorizontalThickLinesPerBand { get; set; } = 200;

        // =========================
        // Backward-compatible API
        // =========================

        /// <summary>
        /// Eski kod tek bir layer bekliyordu.
        /// Set edilirse hem ince hem kalın layer buna eşitlenir.
        /// Get: kalın layer döner.
        /// </summary>
        public string GridLayer
        {
            get { return GridThickLayer; }
            set
            {
                if (!string.IsNullOrWhiteSpace(value))
                {
                    GridThinLayer = value;
                    GridThickLayer = value;
                }
            }
        }

        /// <summary>
        /// Eski kod: VerticalGridStepMeters (metre).
        /// Yeni: VerticalStepMeters
        /// </summary>
        public double VerticalGridStepMeters
        {
            get { return VerticalStepMeters; }
            set { VerticalStepMeters = value; }
        }

        /// <summary>
        /// Eski kod: HorizontalGridStepMeters (metre) - tek adım.
        /// Biz bunu "ince" adım olarak map ediyoruz.
        /// Kalın grid ayrıca 50m ile çizilir.
        /// </summary>
        public double HorizontalGridStepMeters
        {
            get { return HorizontalThinStepMeters; }
            set { HorizontalThinStepMeters = value; }
        }

        /// <summary>
        /// Eski kod: MaxVerticalLines
        /// </summary>
        public int MaxVerticalLines
        {
            get { return MaxVerticalLinesPerBand; }
            set { MaxVerticalLinesPerBand = value; }
        }

        /// <summary>
        /// Eski kod: MaxHorizontalLines (tek limit).
        /// Biz bunu "ince" yatay limitine map ediyoruz.
        /// </summary>
        public int MaxHorizontalLines
        {
            get { return MaxHorizontalThinLinesPerBand; }
            set { MaxHorizontalThinLinesPerBand = value; }
        }

        public CadProfileGridPrinter(AcadDocument doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
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

        public void DrawBandGrid(
            int bandIndex,
            int topLevelMeters,
            int baseLevelMeters,
            double startMeters,
            double endMeters,
            CadProfileTransformer transformer)
        {
            if (transformer == null) throw new ArgumentNullException(nameof(transformer));
            if (endMeters <= startMeters) return;
            if (topLevelMeters <= baseLevelMeters) return;

            AcadModelSpace ms = _doc.ModelSpace;

            // Band CAD sınırları
            double xLeft = transformer.ToCadX(startMeters, bandIndex);
            double xRight = transformer.ToCadX(endMeters, bandIndex);

            double yTop = transformer.ToCadY(topLevelMeters, bandIndex);
            double yBottom = transformer.ToCadY(baseLevelMeters, bandIndex);

            double yMin = Math.Min(yTop, yBottom);
            double yMax = Math.Max(yTop, yBottom);

            // 1) Border (kalın)
            AddLine(ms, xLeft, yMin, xRight, yMin, GridThickLayer);
            AddLine(ms, xRight, yMin, xRight, yMax, GridThickLayer);
            AddLine(ms, xRight, yMax, xLeft, yMax, GridThickLayer);
            AddLine(ms, xLeft, yMax, xLeft, yMin, GridThickLayer);

            // 2) Yatay ince (array)
            DrawHorizontalArray(
                ms, bandIndex,
                baseLevelMeters, topLevelMeters,
                SafeStep(HorizontalThinStepMeters, 1.0),
                xLeft, xRight,
                GridThinLayer,
                MaxHorizontalThinLinesPerBand,
                transformer);

            // 3) Yatay kalın (array)
            DrawHorizontalArray(
                ms, bandIndex,
                baseLevelMeters, topLevelMeters,
                SafeStep(HorizontalThickStepMeters, 5.0),
                xLeft, xRight,
                GridThickLayer,
                MaxHorizontalThickLinesPerBand,
                transformer);

            // 4) Dikey gridler - özel mantık: ilk ve son kalın, diğerleri KmÇizgisi
            DrawVerticalGridsCustom(
                ms, bandIndex,
                startMeters, endMeters,
                SafeStep(VerticalStepMeters, 100.0),
                yMin, yMax,
                MaxVerticalLinesPerBand,
                transformer);
        }

        private static double SafeStep(double step, double fallback)
        {
            if (double.IsNaN(step) || double.IsInfinity(step) || step <= 0) return fallback;
            return step;
        }

        private void DrawVerticalGridsCustom(
            AcadModelSpace ms,
            int bandIndex,
            double startMeters,
            double endMeters,
            double stepMeters,
            double yMin,
            double yMax,
            int maxCount,
            CadProfileTransformer transformer)
        {
            // Global tick hesapla
            double firstTick = NextMultiple(startMeters, stepMeters);
            if (firstTick >= endMeters) return;

            // Tüm km değerlerini hesapla
            var kmValues = new System.Collections.Generic.List<double>();
            double currentKm = firstTick;
            while (currentKm < endMeters && kmValues.Count < maxCount)
            {
                kmValues.Add(currentKm);
                currentKm += stepMeters;
            }

            if (kmValues.Count == 0) return;

            // Her bir dikey çizgiyi ayrı ayrı çiz
            for (int i = 0; i < kmValues.Count; i++)
            {
                double km = kmValues[i];
                double xCad = transformer.ToCadX(km, bandIndex);

                // İlk ve son çizgi Grid-kalın, diğerleri VerticalGridLayer (KmÇizgisi)
                string layerToUse = (i == 0 || i == kmValues.Count - 1) 
                    ? GridThickLayer 
                    : VerticalGridLayer;

                AcadLine ln = null;
                try
                {
                    ln = (AcadLine)ms.AddLine(
                        new double[] { xCad, yMin, 0.0 },
                        new double[] { xCad, yMax, 0.0 });
                    ln.Layer = layerToUse;
                }
                finally
                {
                    ReleaseCom(ln);
                }
            }
        }

        private void DrawVerticalArray(
            AcadModelSpace ms,
            int bandIndex,
            double startMeters,
            double endMeters,
            double stepMeters,
            double yMin,
            double yMax,
            string layer,
            int maxCount,
            CadProfileTransformer transformer)
        {
            // Global tick
            double firstTick = NextMultiple(startMeters, stepMeters);
            if (firstTick >= endMeters) return;

            double xFirst = transformer.ToCadX(firstTick, bandIndex);
            double xNext = transformer.ToCadX(firstTick + stepMeters, bandIndex);

            double colSpacing = Math.Abs(xNext - xFirst);
            if (colSpacing <= 0.000001) return;

            double xEnd = transformer.ToCadX(endMeters, bandIndex);
            double width = Math.Abs(xEnd - xFirst);
            int cols = (int)Math.Floor(width / colSpacing) + 1;
            if (cols <= 0) return;

            if (maxCount > 0 && cols > maxCount)
            {
                int factor = (int)Math.Ceiling(cols / (double)maxCount);
                if (factor < 2) factor = 2;

                stepMeters = stepMeters * factor;
                firstTick = NextMultiple(startMeters, stepMeters);
                if (firstTick >= endMeters) return;

                xFirst = transformer.ToCadX(firstTick, bandIndex);
                xNext = transformer.ToCadX(firstTick + stepMeters, bandIndex);

                colSpacing = Math.Abs(xNext - xFirst);
                if (colSpacing <= 0.000001) return;

                xEnd = transformer.ToCadX(endMeters, bandIndex);
                width = Math.Abs(xEnd - xFirst);
                cols = (int)Math.Floor(width / colSpacing) + 1;
                if (cols <= 0) return;

                if (cols > maxCount) cols = maxCount;
            }

            AcadLine ln = null;
            try
            {
                ln = (AcadLine)ms.AddLine(
                    new double[] { xFirst, yMin, 0.0 },
                    new double[] { xFirst, yMax, 0.0 });

                ln.Layer = layer;

                // 1 row, N cols
                ln.ArrayRectangular(1, cols, 1, 1.0, colSpacing, 1.0);
            }
            finally
            {
                ReleaseCom(ln);
            }
        }

        private void DrawHorizontalArray(
            AcadModelSpace ms,
            int bandIndex,
            double baseLevelMeters,
            double topLevelMeters,
            double stepMeters,
            double xLeft,
            double xRight,
            string layer,
            int maxCount,
            CadProfileTransformer transformer)
        {
            double firstTick = NextMultiple(baseLevelMeters, stepMeters);
            if (firstTick >= topLevelMeters) return;

            double yFirst = transformer.ToCadY(firstTick, bandIndex);
            double yNext = transformer.ToCadY(firstTick + stepMeters, bandIndex);

            double rowSpacing = Math.Abs(yNext - yFirst);
            if (rowSpacing <= 0.000001) return;

            double yEnd = transformer.ToCadY(topLevelMeters, bandIndex);
            double height = Math.Abs(yEnd - yFirst);
            int rows = (int)Math.Floor(height / rowSpacing) + 1;
            if (rows <= 0) return;

            if (maxCount > 0 && rows > maxCount)
            {
                int factor = (int)Math.Ceiling(rows / (double)maxCount);
                if (factor < 2) factor = 2;

                stepMeters = stepMeters * factor;
                firstTick = NextMultiple(baseLevelMeters, stepMeters);
                if (firstTick >= topLevelMeters) return;

                yFirst = transformer.ToCadY(firstTick, bandIndex);
                yNext = transformer.ToCadY(firstTick + stepMeters, bandIndex);

                rowSpacing = Math.Abs(yNext - yFirst);
                if (rowSpacing <= 0.000001) return;

                yEnd = transformer.ToCadY(topLevelMeters, bandIndex);
                height = Math.Abs(yEnd - yFirst);
                rows = (int)Math.Floor(height / rowSpacing) + 1;
                if (rows <= 0) return;

                if (rows > maxCount) rows = maxCount;
            }

            AcadLine ln = null;
            try
            {
                ln = (AcadLine)ms.AddLine(
                    new double[] { xLeft, yFirst, 0.0 },
                    new double[] { xRight, yFirst, 0.0 });

                ln.Layer = layer;

                // N rows, 1 col
                ln.ArrayRectangular(rows, 1, 1, rowSpacing, 1.0, 1.0);
            }
            finally
            {
                ReleaseCom(ln);
            }
        }

        private static double NextMultiple(double value, double step)
        {
            if (step <= 0) return value;
            double k = Math.Ceiling(value / step);
            return k * step;
        }

        private static void AddLine(AcadModelSpace ms, double x1, double y1, double x2, double y2, string layer)
        {
            AcadLine ln = null;
            try
            {
                ln = (AcadLine)ms.AddLine(
                    new double[] { x1, y1, 0.0 },
                    new double[] { x2, y2, 0.0 });

                ln.Layer = layer;
            }
            finally
            {
                ReleaseCom(ln);
            }
        }

        private static void ReleaseCom(object o)
        {
            try
            {
                if (o != null && Marshal.IsComObject(o))
                    Marshal.ReleaseComObject(o);
            }
            catch { }
        }
    }
}
