using Autodesk.AutoCAD.Interop;                 // AcadDocument
using Autodesk.AutoCAD.Interop.Common;          // alignment enums
using PlanProfilYeni.Domain;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace PlanProfilYeni.Cad
{
    public sealed class CadHydraulicTablePrinter
    {
        private readonly AcadDocument _doc;
        private readonly CadHydraulicTableOptions _opt;

        public CadHydraulicTablePrinter(AcadDocument doc, CadHydraulicTableOptions opt)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (opt == null) throw new ArgumentNullException(nameof(opt));

            if (opt.ColumnBoundaries == null || opt.ColumnBoundaries.Length < 2)
                throw new ArgumentException("ColumnBoundaries en az 2 değer içermeli.", nameof(opt));

            _doc = doc;
            _opt = opt;
        }

        public void Print(IReadOnlyList<HydraulicRow> rows)
        {
            if (rows == null) throw new ArgumentNullException(nameof(rows));

            EnsureLayer(_opt.GridLayer);
            EnsureLayer(_opt.TextLayer);

            // 1) Header block
            var basePt = Pt(_opt.BaseX, _opt.BaseY, _opt.BaseZ);
            _doc.ModelSpace.InsertBlock(basePt, _opt.HeaderBlockName, 1.0, 1.0, 1.0, 0.0);

            // 2) Column centers
            var centers = ComputeColumnCenters(_opt.ColumnBoundaries);

            // 3) Text rows + horizontal grid
            double currentY = _opt.BaseY - _opt.FirstRowOffsetY;

            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                var cells = MapRowToCells(r);

                int colCount = Math.Min(cells.Length, centers.Length);
                for (int col = 0; col < colCount; col++)
                {
                    var p = Pt(centers[col], currentY, 0.0);
                    AddCenteredText(cells[col], p);
                }

                var yLine = _opt.BaseY - (i + 1) * _opt.RowStep;
                AddLine(
                    Pt(_opt.BaseX, yLine, 0.0),
                    Pt(_opt.BaseX + _opt.TableWidth, yLine, 0.0),
                    _opt.GridLayer
                );

                currentY -= _opt.RowStep;
            }

            // 4) Vertical grid
            double bottomY = _opt.BaseY - (rows.Count * _opt.RowStep);

            for (int i = 0; i < _opt.ColumnBoundaries.Length; i++)
            {
                double x = _opt.ColumnBoundaries[i];
                AddLine(
                    Pt(x, _opt.BaseY, 0.0),
                    Pt(x, bottomY, 0.0),
                    _opt.GridLayer
                );
            }
        }

        private string[] MapRowToCells(HydraulicRow r)
        {
            var ci = CultureInfo.InvariantCulture;

            string km = r.Kilometre ?? "-";
            bool hasHydrantDetail = r.HizmetAlan.HasValue;

            string hizmetAlan = hasHydrantDetail ? F(r.HizmetAlan, "0.00", ci) : "-";
            string debi = hasHydrantDetail ? F(r.Debi, "0.00", ci) : "-";
            string cikis = hasHydrantDetail ? I(r.CikisSayisi, ci) : "-";
            string tipi = hasHydrantDetail ? (r.Tipi ?? "-") : "-";
            string dyn = hasHydrantDetail ? F(r.DinamikBasinc, "0.00", ci) : "-";

            // Domain’de hesapladık: VAR/YOK/-
            string reg = r.BasincRegulatoru.HasValue ? (r.BasincRegulatoru.Value ? "VAR" : "YOK") : "-";
            string lim = r.DebiLimitoru.HasValue ? (r.DebiLimitoru.Value ? "VAR" : "YOK") : "-";

            return new[]
            {
                r.HatAdi ?? "-",
                r.KesitNo.ToString(ci),
                km.Replace(" - ", " ~ "),
                F(r.BrutAlan, "0.00", ci),
                F(r.SulamaModulu, "0.00", ci),
                F(r.HatDebi, "0.00", ci),
                F(r.Hiz, "0.00", ci),
                F(r.HidrolikEğim, "0.00000", ci),
                r.BoruCinsi ?? "-",
                F(r.IcCap, "0.##", ci),
                F(r.DisCap, "0.##", ci),
                F(r.StatikBasinc, "0.00", ci),
                F(r.DayanmaBasinci, "0.##", ci),
                r.HidrantNo ?? "-",
                hizmetAlan,
                debi,
                cikis,
                tipi,
                dyn,
                reg,
                lim
            };
        }

        private void AddCenteredText(string text, double[] insertPoint)
        {
            if (string.IsNullOrEmpty(text)) text = "-";

            var t = _doc.ModelSpace.AddText(text, insertPoint, _opt.TextHeight);
            t.Layer = _opt.TextLayer;

            // Alignment: COM’da doğru sonuç için hem Alignment hem TextAlignmentPoint set edilir
            t.Alignment = _opt.TextAlignmentCenter;
            t.TextAlignmentPoint = insertPoint;

            // Ek olarak horizontal/vertical alignment sabitlemek daha stabil:
            t.Alignment = _opt.TextAlignmentCenter;
            t.TextAlignmentPoint = insertPoint;


            ComRelease.Release(t);
        }

        private void AddLine(double[] start, double[] end, string layer)
        {
            var ln = _doc.ModelSpace.AddLine(start, end);
            ln.Layer = layer;
            ComRelease.Release(ln);
        }

        private void EnsureLayer(string layerName)
        {
            if (string.IsNullOrWhiteSpace(layerName)) return;

            try
            {
                var layer = _doc.Layers.Item(layerName);
                ComRelease.Release(layer);
            }
            catch
            {
                var created = _doc.Layers.Add(layerName);
                ComRelease.Release(created);
            }
        }

        private static double[] Pt(double x, double y, double z)
        {
            return new[] { x, y, z };
        }

        private static double[] ComputeColumnCenters(double[] boundaries)
        {
            int colCount = boundaries.Length - 1;
            var centers = new double[colCount];

            for (int i = 0; i < colCount; i++)
                centers[i] = (boundaries[i] + boundaries[i + 1]) / 2.0;

            return centers;
        }

        private static string F(double? v, string fmt, CultureInfo ci)
        {
            return v.HasValue ? v.Value.ToString(fmt, ci) : "-";
        }

        private static string I(int? v, CultureInfo ci)
        {
            return v.HasValue ? v.Value.ToString(ci) : "-";
        }
    }
}
