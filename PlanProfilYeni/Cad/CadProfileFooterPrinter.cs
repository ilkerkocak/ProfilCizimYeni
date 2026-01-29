// Cad/CadProfileFooterPrinter.cs
using System;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;

namespace PlanProfilYeni.Cad
{
    public sealed class CadProfileFooterPrinter
    {
        private readonly AcadDocument _doc;

        public string FooterLayer { get; set; } = "TipKesit-YazıÇizgileri";
        public double TextHeight { get; set; } = 2.0;
        public double KmTextOffsetY { get; set; } = -4.0;
        public double KotTextOffsetX { get; set; } = -6.0;

        public CadProfileFooterPrinter(AcadDocument doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        public void DrawFooter(
            int bandIndex,
            int baseLevel,
            double bandStartKm,   // METRE
            double bandEndKm,     // METRE
            CadProfileTransformer transformer,
            double kmStep)        // METRE
        {
            if (transformer == null) throw new ArgumentNullException(nameof(transformer));

            AcadModelSpace ms = _doc.ModelSpace;

            double stepMeters = (kmStep <= 0) ? 100.0 : kmStep;

            // 1) Chainage yazıları
            double start = Math.Floor(bandStartKm / stepMeters) * stepMeters;

            for (double m = start; m <= bandEndKm + 1e-9; m += stepMeters)
            {
                if (m < bandStartKm - 1e-9) continue;

                double x = transformer.ToCadX(m, bandIndex);
                double y = transformer.ToCadY(baseLevel, bandIndex) + KmTextOffsetY;

                string txtVal = FormatChainageFromMeters(m);

                AcadText txt = null;
                try
                {
                    txt = (AcadText)ms.AddText(
                        txtVal,
                        new double[] { x, y, 0 },
                        TextHeight);

                    txt.Layer = FooterLayer;

                    // ✅ DOĞRU SIRA
                    txt.Alignment = AcAlignment.acAlignmentMiddleCenter;
                    txt.TextAlignmentPoint = new double[] { x, y, 0 };
                }
                finally
                {
                    ReleaseCom(txt);
                }
            }

            // Band sonu özel değer
            double remainder = bandEndKm % stepMeters;
            if (remainder > 1e-6 && (bandEndKm - bandStartKm) > 1e-6)
            {
                double x = transformer.ToCadX(bandEndKm, bandIndex);
                double y = transformer.ToCadY(baseLevel, bandIndex) + KmTextOffsetY;

                AcadText txt = null;
                try
                {
                    txt = (AcadText)ms.AddText(
                        FormatChainageFromMeters(bandEndKm),
                        new double[] { x, y, 0 },
                        TextHeight);

                    txt.Layer = FooterLayer;

                    txt.Alignment = AcAlignment.acAlignmentMiddleCenter;
                    txt.TextAlignmentPoint = new double[] { x, y, 0 };
                }
                finally
                {
                    ReleaseCom(txt);
                }
            }

            // 2) Kot yazıları
            for (int kot = baseLevel; kot <= baseLevel + 13; kot++)
            {
                if (kot % 5 != 0) continue;

                double x = transformer.ToCadX(bandStartKm, bandIndex) + KotTextOffsetX;
                double y = transformer.ToCadY(kot, bandIndex);

                AcadText txt = null;
                try
                {
                    txt = (AcadText)ms.AddText(
                        kot.ToString(),
                        new double[] { x, y, 0 },
                        TextHeight);

                    txt.Layer = FooterLayer;

                    txt.Alignment = AcAlignment.acAlignmentMiddleRight;
                    txt.TextAlignmentPoint = new double[] { x, y, 0 };
                }
                finally
                {
                    ReleaseCom(txt);
                }
            }

            // ❌ ms RELEASE YOK
        }

        private static string FormatChainageFromMeters(double meters)
        {
            if (meters < 0) meters = 0;

            int kmPart = (int)Math.Floor(meters / 1000.0);
            double mPart = meters - (kmPart * 1000.0);

            if (Math.Abs(mPart - Math.Round(mPart)) < 1e-6)
                return string.Format("{0}+{1:000}", kmPart, (int)Math.Round(mPart));

            return string.Format("{0}+{1:000.00}", kmPart, mPart);
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
