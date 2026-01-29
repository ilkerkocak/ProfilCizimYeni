using System;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Interop.Common;
using Microsoft.Office.Interop.Excel;
using PlanProfilYeni.Cad;
using PlanProfilYeni.Domain;
using PlanProfilYeni.Services;

namespace PlanProfilYeni.Application
{
    public sealed class DrawProfileUseCase
    {
        private void EnsureLayer(AcadDocument doc, string layerName, int colorIndex)
        {
            try
            {
                var layer = doc.Layers.Item(layerName);
                layer.color = (AcColor)colorIndex;
            }
            catch
            {
                var layer = doc.Layers.Add(layerName);
                layer.color = (AcColor)colorIndex;
            }
        }

        /// <param name="horizontalScale">
        /// Bu projede "km" aslında METRE.
        /// Bu parametreyi "cadUnitsPerMeter" olarak kullanıyoruz.
        /// Örn: 100m=20 birim => 0.2
        /// </param>
        /// <param name="verticalScale">
        /// Bu projede "km" aslında METRE.
        /// Bu parametreyi de "cadUnitsPerMeter" olarak kullanıyoruz.
        /// Örn: 0.2
        /// </param>
        public void Execute(
            string excelPath,
            AcadDocument doc,
            double insertX,
            double insertY,
            double horizontalScale,
            double verticalScale,
            string lineName)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = null;
            Workbook wb = null;

            try
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                wb = xlApp.Workbooks.Open(excelPath);

                Worksheet wsPressure = wb.Worksheets["Basınç Profili"];
                Worksheet wsProfile = wb.Worksheets["Arazi & Boru Profili"];

                // 1) Arazi & Boru profili
                var profileExcelService = new ProfileExcelService();
                double[] groundXY = profileExcelService.ReadGroundProfile(wsProfile);
                double[] pipeXY = profileExcelService.ReadPipeProfile(wsProfile);

                // 2) Hidrolik seri
                var pressureService = new PressureProfileExcelService();
                double[] hydraulicKmValuePairs = pressureService.ReadHydraulicSeries(wsPressure);

                // 3) Band üretimi
                var bandService = new ProfileBandService();
                ProfileBandSet bands = bandService.BuildBands(groundXY, pipeXY, lineName);

                // 4) CAD dönüşüm (KİLİTLİ)
                // Excel mesafe ekseni: METRE
                // horizontalScale = 0.2 (yatay: 100m = 20 CAD)
                // verticalScale = 10.0 (düşey: 1m = 10 CAD, 1/100 ölçek)
                double cadUnitsPerMeterHorizontal = horizontalScale;
                if (cadUnitsPerMeterHorizontal <= 0) cadUnitsPerMeterHorizontal = 0.2;

                double cadUnitsPerMeterVertical = verticalScale;
                if (cadUnitsPerMeterVertical <= 0) cadUnitsPerMeterVertical = 10.0;

                var transformOptions = new CadProfileTransformOptions
                {
                    OriginX = insertX,
                    GridTopY = insertY,

                    // 🔒 X ekseni: metre tabanlı, yatay ölçek
                    CadUnitsPerKm = horizontalScale,
                    // 🔒 Y ekseni: metre tabanlı, düşey ölçek (verticalScale = 10.0 için 1m=10CAD)
                    CadUnitsPerMeter = verticalScale,

                    BandPanelSpacingX = 0
                };

                var transformer = new CadProfileTransformer(bands, transformOptions);

                // Layerları renklerle oluştur
                // 254 = Açık gri, 150 = Koyu gri, 253 = Gri, 1 = Kırmızı, 34 = Yeşil
                EnsureLayer(doc, "Grid-ince", 254);      // Açık gri (20 cm gridler)
                EnsureLayer(doc, "Grid-kalın", 150);     // Koyu gri (1m gridler + border)
                EnsureLayer(doc, "KmÇizgisi", 253);      // Gri (ara dikey gridler)
                EnsureLayer(doc, "Arazi", 34);            // Yeşil
                EnsureLayer(doc, "Boru", 1);             // Kırmızı
                EnsureLayer(doc, "TipKesit-YazıÇizgileri", 253);  // Gri

                // 5) Grid + Header + Footer (band başına 1 kez)
                var gridPrinter = new CadProfileGridPrinter(doc)
                {
                    GridThinLayer = "Grid-ince",
                    GridThickLayer = "Grid-kalın",
                    VerticalGridLayer = "KmÇizgisi",   // Ara dikey gridler için
                    VerticalStepMeters = 100.0,        // Dikey grid: 100 metre
                    HorizontalThinStepMeters = 0.2,    // İnce yatay grid: 20 cm
                    HorizontalThickStepMeters = 1.0,   // Kalın yatay grid: 1 metre
                    MaxVerticalLines = 400,
                    MaxHorizontalLines = 1500
                };

                var headerPrinter = new CadProfileHeaderPrinter(doc);
                var footerPrinter = new CadProfileFooterPrinter(doc);

                for (int b = 0; b < bands.BandCount; b++)
                {
                    double startM = (b == 0) ? 0.0 : bands.BreakKms[b - 1]; // METRE
                    double endM = bands.BreakKms[b];                         // METRE

                    gridPrinter.DrawBandGrid(
                        b,
                        bands.TopLevels[b],
                        bands.BaseLevels[b],
                        startM,
                        endM,
                        transformer);

                    headerPrinter.DrawHeader(
                        b,
                        startM,
                        endM,
                        transformer,
                        lineName,
                        transformOptions.CadUnitsPerKm,
                        transformOptions.CadUnitsPerMeter);

                    footerPrinter.DrawFooter(
                        b,
                        bands.BaseLevels[b],
                        startM,
                        endM,
                        transformer,
                        100.0);
                }

                // 6) Arazi + Boru
                var clipper = new ProfileClipService();
                var polyDrawer = new CadProfilePolylinePrinter(doc);

                // Arazi çizgisi (siyah, normal kalınlık)
                polyDrawer.DrawSegments(
                    clipper.SplitByBands(groundXY, bands),
                    transformer,
                    "Arazi",
                    0.25);

                // Boru çizgisi (kırmızı, kalın)
                polyDrawer.DrawSegments(
                    clipper.SplitByBands(pipeXY, bands),
                    transformer,
                    "Boru",
                    0.5);

                // 7) Hidrolik
                var hydraulicBuilder = new HydraulicProfileBuilder();
                var hydraulicBuild = hydraulicBuilder.Build(
                    hydraulicKmValuePairs,
                    pipeXY,
                    bands,
                    new HydraulicBuildOptions
                    {
                        VerticalBreakStepMeters = 4.0,
                        ValueMode = HydraulicValueMode.AbsoluteElevationMeters
                    });

                new CadProfileHydraulicDrawer(doc, bands)
                    .Draw(hydraulicBuild.Segments,
                          hydraulicBuild.VerticalBreakMarkers,
                          transformer,
                          transformOptions);

                // 8) Profil ekipmanları
                var hydrants = pressureService.ReadHydrants(wsPressure);
                var hatAyrimlari = pressureService.ReadHatAyrimlari(wsPressure);
                var bkvs = pressureService.ReadBkvs(wsPressure);
                var manualVantuz = pressureService.ReadManualVantuzKms(wsPressure);
                var manualTahliye = pressureService.ReadManualTahliyeKms(wsPressure);

                var extrema = new PipeExtremaDetector();
                extrema.Detect(pipeXY, out var vantuzFromPipe, out var tahliyeFromPipe);

                var equipmentBuilder = new ProfileEquipmentBuilder();
                var equipmentItems = equipmentBuilder.Build(
                    hydrants,
                    hatAyrimlari,
                    bkvs,
                    vantuzFromPipe,
                    tahliyeFromPipe,
                    manualVantuz,
                    manualTahliye);

                new CadProfileEquipmentDrawer(doc, bands, transformOptions)
                    .Draw(equipmentItems, bands, transformer, pipeXY);
            }
            finally
            {
                if (wb != null) { wb.Close(false); ReleaseCom(wb); }
                if (xlApp != null) { xlApp.Quit(); ReleaseCom(xlApp); }
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
