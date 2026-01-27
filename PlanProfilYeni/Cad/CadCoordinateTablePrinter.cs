using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using PlanProfilYeni.Domain;
using System;
using System.Collections.Generic;


namespace PlanProfilYeni.Cad
{
    public class CadCoordinateTablePrinter
    {
        private readonly AcadDocument _doc;

        // Orijinal kolon boundary değerleri (VB kodundaki array3)
        private readonly double[] _columnBoundaries = new double[]
        {
            500.0,
            514.0,
            528.237,
            545.889,
            566.0,
            579.054,
            591.69,
            607.797
        };

        public CadCoordinateTablePrinter(AcadDocument doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        public void Print(List<CoordinateRow> rows)
        {
            if (rows == null || rows.Count == 0)
                throw new ArgumentException("Koordinat satırı yok.");

            // 1) Block insert
            double[] basePt = { 500, -50, 0 };
            _doc.ModelSpace.InsertBlock(basePt, "Koordinat Tablosu", 1, 1, 1, 0);

            // Kolon merkezleri
            var centers = ComputeCenters(_columnBoundaries);

            double y = basePt[1] - 1.75;
            double startX = _columnBoundaries[0];
            double endX = _columnBoundaries[_columnBoundaries.Length - 1];

            // 2) Satırlar
            foreach (var r in rows)
            {
                AddText(r.HatAdi, centers[0], y);
                AddText(r.NoktaAdi, centers[1], y);
                AddText(r.X, centers[2], y);
                AddText(r.Y, centers[3], y);
                AddText(r.Acisal1 ?? "-", centers[4], y);
                AddText(r.Acisal2 ?? "-", centers[5], y);
                AddText(r.Km, centers[6], y);

                // Yatay çizgi
                AddLine(startX, y - 1.75, endX, y - 1.75);

                y -= 3.5;
            }

            // 3) Dikey çizgiler
            double rowHeight = 3.5;
            double top = basePt[1];
            double bottom = basePt[1] - rows.Count * rowHeight;


            foreach (var x in _columnBoundaries)
            {
                AddLine(x, top, x, bottom);
            }
        }

        private void AddText(string text, double x, double y)
        {
            double[] pt = { x, y, 0 };
            var t = _doc.ModelSpace.AddText(text ?? "-", pt, 2.0);
            t.Layer = "TipKesit-YazıÇizgileri";
            t.Alignment = AcAlignment.acAlignmentMiddleCenter;
            t.TextAlignmentPoint = pt;
        }

        private void AddLine(double x1, double y1, double x2, double y2)
        {
            double[] p1 = { x1, y1, 0 };
            double[] p2 = { x2, y2, 0 };
            var ln = _doc.ModelSpace.AddLine(p1, p2);
            ln.Layer = "Arazi";
        }

        private double[] ComputeCenters(double[] boundaries)
        {
            var centers = new double[boundaries.Length - 1];
            for (int i = 0; i < centers.Length; i++)
                centers[i] = (boundaries[i] + boundaries[i + 1]) / 2.0;
            return centers;
        }
    }
}
