// Cad/CadProfileTransformOptions.cs
using System;

namespace PlanProfilYeni.Cad
{
    /// <summary>
    /// Profil dönüşüm parametreleri:
    /// - X ekseni: km -> CAD birimi
    /// - Y ekseni: kot(m) -> CAD birimi
    /// - Origin: band grid'inin sol üst referansı
    /// </summary>
    public sealed class CadProfileTransformOptions
    {
        /// <summary>Grid sol başlangıç X (CAD birimi)</summary>
        public double OriginX { get; set; }

        /// <summary>Grid üst sınır Y (CAD birimi)</summary>
        public double GridTopY { get; set; }

        /// <summary>1 km kaç CAD birimi? (yatay ölçek)</summary>
        public double CadUnitsPerKm { get; set; }

        /// <summary>1 m kot kaç CAD birimi? (düşey ölçek). Class7 görünümünde tipik: 10</summary>
        public double CadUnitsPerMeter { get; set; }

        /// <summary>
        /// Bandlar arası yatay mesafe (pencere yoksa 0). Senin çıktıdaki gibi yan yana panel varsa kullanılır.
        /// Şimdilik 0 bırakabilirsin.
        /// </summary>
        public double BandPanelSpacingX { get; set; }

        public void Validate()
        {
            if (CadUnitsPerKm <= 0) throw new ArgumentOutOfRangeException(nameof(CadUnitsPerKm));
            if (CadUnitsPerMeter <= 0) throw new ArgumentOutOfRangeException(nameof(CadUnitsPerMeter));
        }
    }
}
