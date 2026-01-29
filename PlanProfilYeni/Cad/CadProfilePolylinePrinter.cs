// Cad/CadProfilePolylinePrinter.cs
using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using PlanProfilYeni.Domain;

namespace PlanProfilYeni.Cad
{
    /// <summary>
    /// ProfileClipService çıktısı olan band segmentlerini AddLine ile çizdirir.
    /// Polyline yerine line segmentleri: Class7 çıktısını daha stabil taklit eder.
    /// </summary>
    public sealed class CadProfilePolylinePrinter
    {
        private readonly AcadDocument _doc;

        public double LineWeight { get; set; } = 0.5;  // mm cinsinden çizgi kalınlığı

        public CadProfilePolylinePrinter(AcadDocument doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        public void DrawSegments(
            IEnumerable<ProfilePolylineBandSegment> segments,
            CadProfileTransformer transformer,
            string layerName,
            double? lineWeight = null)
        {
            if (segments == null) throw new ArgumentNullException(nameof(segments));
            if (transformer == null) throw new ArgumentNullException(nameof(transformer));
            if (string.IsNullOrWhiteSpace(layerName)) throw new ArgumentException("Layer name required.", nameof(layerName));

            double weight = lineWeight ?? LineWeight;
            AcadModelSpace ms = _doc.ModelSpace;

            foreach (var seg in segments)
            {
                var pairs = seg.KmElevationPairs;
                int bandIndex = seg.BandIndex;

                // pairs: [km0,kot0, km1,kot1, ...]
                for (int i = 2; i < pairs.Length; i += 2)
                {
                    double km1 = pairs[i - 2];
                    double y1 = pairs[i - 1];
                    double km2 = pairs[i];
                    double y2 = pairs[i + 1];

                    double[] p1 = transformer.ToCadPoint(km1, y1, bandIndex);
                    double[] p2 = transformer.ToCadPoint(km2, y2, bandIndex);

                    AcadLine ln = null;
                    try
                    {
                        ln = (AcadLine)ms.AddLine(p1, p2);
                        ln.Layer = layerName;
                        
                        // Çizgi kalınlığı ayarla (mm cinsinden * 100)
                        // 0.25mm = 25, 0.50mm = 50 gibi
                        int lwValue = (int)(weight * 100);
                        ln.Lineweight = (ACAD_LWEIGHT)lwValue;
                    }
                    finally
                    {

                    }
                }
            }

        
        }
    }
}
