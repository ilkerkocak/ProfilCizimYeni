// Domain/ProfilePolylineBandSegment.cs
using System;

namespace PlanProfilYeni.Domain
{
    /// <summary>
    /// Bir polylinenin, band kırıklarına göre elde edilen tek bir band içindeki parçası.
    /// Noktalar (km, kot) uzayındadır: [km0, kot0, km1, kot1, ...]
    /// </summary>
    public sealed class ProfilePolylineBandSegment
    {
        public int BandIndex { get; private set; }
        public double[] KmElevationPairs { get; private set; } // [x,y] pairs

        public double StartKm { get { return KmElevationPairs != null && KmElevationPairs.Length >= 2 ? KmElevationPairs[0] : 0.0; } }
        public double EndKm { get { return KmElevationPairs != null && KmElevationPairs.Length >= 2 ? KmElevationPairs[KmElevationPairs.Length - 2] : 0.0; } }

        public ProfilePolylineBandSegment(int bandIndex, double[] kmElevationPairs)
        {
            if (kmElevationPairs == null) throw new ArgumentNullException(nameof(kmElevationPairs));
            if (kmElevationPairs.Length < 4 || (kmElevationPairs.Length % 2) != 0)
                throw new ArgumentException("KmElevationPairs must be [km,kot] pairs, length>=4 and even.", nameof(kmElevationPairs));

            BandIndex = bandIndex;
            KmElevationPairs = kmElevationPairs;
        }
    }
}
