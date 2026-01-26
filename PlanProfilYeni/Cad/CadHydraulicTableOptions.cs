using Autodesk.AutoCAD.Interop.Common;

namespace PlanProfilYeni.Cad
{
    public sealed class CadHydraulicTableOptions
    {
        public double BaseX { get; set; }
        public double BaseY { get; set; }
        public double BaseZ { get; set; }

        public string HeaderBlockName { get; set; }

        // Column boundaries (X positions) - length = columnCount + 1
        public double[] ColumnBoundaries { get; set; }

        public double TableWidth { get; set; }
        public double RowStep { get; set; }
        public double FirstRowOffsetY { get; set; }

        public string GridLayer { get; set; }
        public string TextLayer { get; set; }

        public double TextHeight { get; set; }

        // COM alignment:
        public AcAlignment TextAlignmentCenter { get; set; }
    }
}
