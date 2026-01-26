namespace PlanProfilYeni.Application
{
    public sealed class HydraulicProcessOptions
    {
        public string[] InputFiles { get; set; }

        // Excel export
        public string OutputExcelPath { get; set; }
        public bool OpenExcelAfterExport { get; set; }

        // CAD / çizim opsiyonları (şimdilik burada dursun)
        public bool DrawBlocks { get; set; }
        public int VerticalScale { get; set; }
        public int HorizontalScale { get; set; }
    }
}
