// Domain/HydraulicDrawableSegment.cs
using System;

namespace PlanProfilYeni.Domain
{
    /// <summary>
    /// Hidrolik çizgi parçası.
    /// BandIndex: band km kırığına göre.
    /// TopRefMeters: band içi mikro kırık için topLimit override (Class7 topLimit -= 4.0).
    /// Points: (km, elevationMeters) çiftleri.
    /// </summary>
    public sealed class HydraulicDrawableSegment
    {
        public int BandIndex { get; private set; }
        public double TopRefMeters { get; private set; }
        public double[] KmElevationPairs { get; private set; }
        public double BandTopLevelMeters { get; set; }


        public HydraulicDrawableSegment(int bandIndex, double topRefMeters, double[] kmElevationPairs)
        {
            if (kmElevationPairs == null) throw new ArgumentNullException(nameof(kmElevationPairs));
            if (kmElevationPairs.Length < 4 || (kmElevationPairs.Length % 2) != 0)
                throw new ArgumentException("KmElevationPairs must be [km,elev] pairs, length>=4 and even.", nameof(kmElevationPairs));

            BandIndex = bandIndex;
            TopRefMeters = topRefMeters;
            KmElevationPairs = kmElevationPairs;
        }
    }
}
