// Cad/CadProfileHeaderPrinter.cs
using System;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;

namespace PlanProfilYeni.Cad
{
    /// <summary>
    /// Profil üst bilgileri (başlık).
    /// Her band için TEKRAR basılır.
    /// </summary>
    public sealed class CadProfileHeaderPrinter
    {
        private readonly AcadDocument _doc;

        public string HeaderLayer { get; set; } = "TipKesit-YazıÇizgileri";
        public double TextHeight { get; set; } = 2.5;
        public double HeaderOffsetY { get; set; } = 6.0;

        public CadProfileHeaderPrinter(AcadDocument doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        public void DrawHeader(
            int bandIndex,
            double bandStartKm,
            double bandEndKm,
            CadProfileTransformer transformer,
            string lineName,
            double horizontalScale,
            double verticalScale)
        {
            AcadModelSpace ms = _doc.ModelSpace;

            double xMid = transformer.ToCadX((bandStartKm + bandEndKm) / 2.0, bandIndex);
            double y = transformer.ToCadY(transformer
                .GetType() != null ? transformer.ToCadY(0, bandIndex) : 0, bandIndex); // dummy safe

            // GridTopY + offset
            y = transformer.ToCadY(transformer
                .GetType() != null ? 0 : 0, bandIndex) + HeaderOffsetY;

            // Başlık metni
            string text =
                $"{lineName}  |  Yatay Ölçek: 1/{horizontalScale:0}  Düşey Ölçek: 1/{verticalScale:0}";

            AcadText txt = null;
            try
            {
                txt = (AcadText)ms.AddText(text,
                    new double[] { xMid, y, 0 },
                    TextHeight);
                txt.Alignment = AcAlignment.acAlignmentMiddleCenter;
                txt.TextAlignmentPoint = new double[] { xMid, y, 0 };
                txt.Layer = HeaderLayer;
            }
            finally
            {

            }
        }
    }
}
