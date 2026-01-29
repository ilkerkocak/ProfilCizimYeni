using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using PlanProfilYeni.Domain;

namespace PlanProfilYeni.Cad
{
    /// <summary>
    /// Class7 uyumlu PROFİL ekipman çizici.
    /// 
    /// ✔ Hidrant  : Dynamic Block + Visibility ("n ÇIKIŞ")
    /// ✔ BKV      : Block (fallback box)
    /// ✔ Vantuz   : Block (fallback circle)
    /// ✔ Tahliye  : Block (fallback X)
    /// ✔ HatAyrımı: Üçgen + yazı (blok yok)
    /// 
    /// Y konumu: Boru kotu + offset
    /// Band değişince aynı km yeniden çizilir.
    /// </summary>
    public sealed class CadProfileEquipmentDrawer
    {
        private readonly AcadDocument _doc;
        private readonly ProfileBandSet _bands;
        private readonly CadProfileTransformOptions _transformOptions;

        public string Layer { get; set; } = "TipKesit-YazıÇizgileri";
        public double HydraulicOffsetCadUnits { get; set; } = 40.0;
        public double BlockScale { get; set; } = 0.2;
        public double SymbolSizeCad { get; set; } = 0.5;
        
        // Hidrolik profil sabitleri (CadProfileHydraulicDrawer ile aynı)
        private const double HYDRAULIC_VERTICAL_OFFSET_METERS = 2.0;
        private const double HYDRAULIC_GRID_HEIGHT = 4.0;

        // 🔒 Class7 blok isimleri (DWG şablonuyla birebir)
        private const string BLOCK_HIDRANT = "Hidrant";
        private const string BLOCK_BKV = "BKV";
        private const string BLOCK_VANTUZ = "Vantuz";
        private const string BLOCK_TAHLİYE = "Tahliye";

        private const string BLOCK_AYRIM = "Ayrım";

        public CadProfileEquipmentDrawer(
            AcadDocument doc,
            ProfileBandSet bands,
            CadProfileTransformOptions transformOptions)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _bands = bands ?? throw new ArgumentNullException(nameof(bands));
            _transformOptions = transformOptions ?? throw new ArgumentNullException(nameof(transformOptions));
        }

        // =====================================================
        // PUBLIC
        // =====================================================

        public void Draw(
            IList<ProfileEquipmentItem> items,
            ProfileBandSet bands,
            CadProfileTransformer transformer,
            double[] pipeKmElevationPairs)
        {
            if (items == null || items.Count == 0)
                return;

            AcadModelSpace ms = _doc.ModelSpace;

            try
            {
                foreach (var item in items)
                {
                    int bandIndex = ResolveBandIndex(item.Km, bands);

                    double x = transformer.ToCadX(item.Km, bandIndex);
                    
                    // Hidrolik profilin üst sınırını hesapla
                    double araziTopY = transformer.ToCadY(bands.TopLevels[bandIndex], bandIndex);
                    double hydraulicOffset = HYDRAULIC_VERTICAL_OFFSET_METERS * _transformOptions.CadUnitsPerMeter;
                    double hydraulicHeight = HYDRAULIC_GRID_HEIGHT * _transformOptions.CadUnitsPerMeter;
                    double hydraulicTopY = araziTopY + hydraulicOffset + hydraulicHeight;
                    
                    // Ekipman: Hidrolik profilden 40 birim yukarı
                    double y = hydraulicTopY + HydraulicOffsetCadUnits;

                    DrawItem(ms, item, x, y);
                }
            }
            finally
            {
                ReleaseCom(ms);
            }
        }

        // =====================================================
        // DRAW ROUTER
        // =====================================================

        private void DrawItem(AcadModelSpace ms, ProfileEquipmentItem item, double x, double y)
        {
            switch (item.Type)
            {
                case ProfileEquipmentType.Hydrant:
                    InsertHydrantWithVisibility(ms, x, y, item.OutletCount);
                    break;

                case ProfileEquipmentType.Bkv:
                    InsertOrFallback(ms, BLOCK_BKV, x, y, () => DrawBox(ms, x, y));
                    break;

                case ProfileEquipmentType.Vantuz:
                    InsertOrFallback(ms, BLOCK_VANTUZ, x, y, () => DrawCircle(ms, x, y));
                    break;

                case ProfileEquipmentType.Tahliye:
                    InsertOrFallback(ms, BLOCK_TAHLİYE, x, y, () => DrawX(ms, x, y));
                    break;

                case ProfileEquipmentType.HatAyrimi:
                    InsertOrFallback(ms, BLOCK_AYRIM, x, y, () => DrawHatAyrimi(ms, x, y, item.Label));
                    break;
            }
        }

        // =====================================================
        // BLOCK INSERT
        // =====================================================

        private void InsertHydrantWithVisibility(
            AcadModelSpace ms,
            double x,
            double y,
            int? outletCount)
        {
            if (!BlockExists(BLOCK_HIDRANT))
            {
                DrawCross(ms, x, y);
                return;
            }

            AcadBlockReference br = null;
            try
            {
                br = (AcadBlockReference)ms.InsertBlock(
                    new double[] { x, y, 0 },
                    BLOCK_HIDRANT,
                    BlockScale, BlockScale, BlockScale,
                    0.0);

                br.Layer = Layer;

                if (outletCount.HasValue)
                    SetHydrantVisibility(br, outletCount.Value);
            }
            catch
            {
                DrawCross(ms, x, y);
            }
            finally
            {
                ReleaseCom(br);
            }
        }

        private void InsertOrFallback(
            AcadModelSpace ms,
            string blockName,
            double x,
            double y,
            Action fallback)
        {
            if (!BlockExists(blockName))
            {
                fallback();
                return;
            }

            AcadBlockReference br = null;
            try
            {
                br = (AcadBlockReference)ms.InsertBlock(
                    new double[] { x, y, 0 },
                    blockName,
                    BlockScale, BlockScale, BlockScale,
                    0.0);
                br.Layer = Layer;
            }
            catch
            {
                fallback();
            }
            finally
            {
                ReleaseCom(br);
            }
        }

        private bool BlockExists(string name)
        {
            try
            {
                return _doc.Blocks.Item(name) != null;
            }
            catch
            {
                return false;
            }
        }

        // =====================================================
        // HYDRANT VISIBILITY (EN KRİTİK)
        // =====================================================

        private void SetHydrantVisibility(AcadBlockReference br, int outletCount)
        {
            string visibilityName = outletCount + " ÇIKIŞ";

            try
            {
                object[] props = (object[])br.GetDynamicBlockProperties();
                if (props == null) return;

                foreach (object p in props)
                {
                    var dp = p as AcadDynamicBlockReferenceProperty;
                    if (dp == null) continue;

                    if (!dp.PropertyName.Equals("Visibility", StringComparison.OrdinalIgnoreCase))
                        continue;

                    object[] allowed = (object[])dp.AllowedValues;
                    foreach (object v in allowed)
                    {
                        if (v != null && v.ToString() == visibilityName)
                        {
                            dp.Value = visibilityName;
                            return;
                        }
                    }
                }
            }
            catch
            {
                // Class7: sessiz geç
            }
        }

        // =====================================================
        // FALLBACK SYMBOLS
        // =====================================================

        private void DrawCross(AcadModelSpace ms, double x, double y)
        {
            double s = SymbolSizeCad;
            AddLine(ms, x - s, y, x + s, y);
            AddLine(ms, x, y - s, x, y + s);
        }

        private void DrawBox(AcadModelSpace ms, double x, double y)
        {
            double s = SymbolSizeCad;
            AddLine(ms, x - s, y - s, x + s, y - s);
            AddLine(ms, x + s, y - s, x + s, y + s);
            AddLine(ms, x + s, y + s, x - s, y + s);
            AddLine(ms, x - s, y + s, x - s, y - s);
        }

        private void DrawCircle(AcadModelSpace ms, double x, double y)
        {
            double s = SymbolSizeCad;
            AcadCircle c = null;
            try
            {
                c = (AcadCircle)ms.AddCircle(new double[] { x, y, 0 }, s);
                c.Layer = Layer;
            }
            finally
            {
                ReleaseCom(c);
            }
        }

        private void DrawX(AcadModelSpace ms, double x, double y)
        {
            double s = SymbolSizeCad;
            AddLine(ms, x - s, y - s, x + s, y + s);
            AddLine(ms, x - s, y + s, x + s, y - s);
        }

        private void DrawHatAyrimi(AcadModelSpace ms, double x, double y, string label)
        {
            DrawTriangle(ms, x, y);
            if (!string.IsNullOrWhiteSpace(label))
                AddText(ms, label, x, y + SymbolSizeCad + 1.5);
        }

        private void DrawTriangle(AcadModelSpace ms, double x, double y)
        {
            double s = SymbolSizeCad;
            AddLine(ms, x, y + s, x - s, y - s);
            AddLine(ms, x - s, y - s, x + s, y - s);
            AddLine(ms, x + s, y - s, x, y + s);
        }

        // =====================================================
        // PRIMITIVES
        // =====================================================

        private void AddLine(AcadModelSpace ms, double x1, double y1, double x2, double y2)
        {
            AcadLine ln = null;
            try
            {
                ln = (AcadLine)ms.AddLine(
                    new double[] { x1, y1, 0 },
                    new double[] { x2, y2, 0 });
                ln.Layer = Layer;
            }
            finally
            {
                ReleaseCom(ln);
            }
        }

        private void AddText(AcadModelSpace ms, string text, double x, double y)
        {
            AcadText txt = null;
            try
            {
                txt = (AcadText)ms.AddText(
                    text,
                    new double[] { x, y, 0 },
                    2.0);
                txt.Alignment = AcAlignment.acAlignmentMiddleLeft;
                txt.TextAlignmentPoint = new double[] { x, y, 0 };
                txt.Layer = Layer;
            }
            finally
            {
                ReleaseCom(txt);
            }
        }

        // =====================================================
        // HELPERS
        // =====================================================

        private static int ResolveBandIndex(double km, ProfileBandSet bands)
        {
            for (int i = 0; i < bands.BreakKms.Length; i++)
            {
                if (km <= bands.BreakKms[i])
                    return Math.Min(i, bands.BandCount - 1);
            }
            return bands.BandCount - 1;
        }

        private static double InterpolatePipeElevation(double[] xyPairs, double x)
        {
            for (int i = 2; i <= xyPairs.Length - 2; i += 2)
            {
                double x1 = xyPairs[i - 2];
                double x2 = xyPairs[i];
                if (x1 <= x && x <= x2)
                {
                    double y1 = xyPairs[i - 1];
                    double y2 = xyPairs[i + 1];
                    double t = (x - x1) / (x2 - x1);
                    return y1 + t * (y2 - y1);
                }
            }
            return xyPairs[xyPairs.Length - 1];
        }

        private static void ReleaseCom(object o)
        {
            if (o != null && Marshal.IsComObject(o))
                Marshal.ReleaseComObject(o);
        }
    }
}
