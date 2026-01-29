// Cad/CadProfileTransformer.cs
using System;
using PlanProfilYeni.Domain;

namespace PlanProfilYeni.Cad
{
    /// <summary>
    /// (km, kot) -> CAD (x,y) dönüşümü.
    /// Class7 mantığı: Y dönüşümünde bandın TopLevel'i referans alınır:
    /// yCad = GridTopY - (TopLevel - elevation) * CadUnitsPerMeter
    /// </summary>
    public sealed class CadProfileTransformer
    {
        private readonly ProfileBandSet _bands;
        private readonly CadProfileTransformOptions _opt;

        public CadProfileTransformer(ProfileBandSet bands, CadProfileTransformOptions options)
        {
            _bands = bands ?? throw new ArgumentNullException(nameof(bands));
            _opt = options ?? throw new ArgumentNullException(nameof(options));
            _opt.Validate();
        }

        public double ToCadX(double km, int bandIndex)
        {
            // Eğer bandları yatay paneller şeklinde yan yana basıyorsan bandIndex ile offset ver.
            double panelOffset = (bandIndex > 0) ? bandIndex * _opt.BandPanelSpacingX : 0.0;
            return _opt.OriginX + panelOffset + km * _opt.CadUnitsPerKm;
        }

        public double ToCadY(double elevationMeters, int bandIndex)
        {
            if (bandIndex < 0) bandIndex = 0;
            if (bandIndex >= _bands.BandCount) bandIndex = _bands.BandCount - 1;

            int topLevel = _bands.TopLevels[bandIndex];

            // GridTopY üst sınır; kot aşağı indikçe y azalmalı (AutoCAD Y yukarı pozitif).
            // Class7: y = baseY - (topLevel - value) * scale
            return _opt.GridTopY - (topLevel - elevationMeters) * _opt.CadUnitsPerMeter;
        }

        public double[] MakePoint(double xCad, double yCad)
        {
            return new double[3] { xCad, yCad, 0.0 };
        }

        public double[] ToCadPoint(double km, double elevationMeters, int bandIndex)
        {
            double x = ToCadX(km, bandIndex);
            double y = ToCadY(elevationMeters, bandIndex);
            return new double[3] { x, y, 0.0 };
        }
    }
}
