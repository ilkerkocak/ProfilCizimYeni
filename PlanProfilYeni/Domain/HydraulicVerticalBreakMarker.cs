// Domain/HydraulicVerticalBreakMarker.cs
namespace PlanProfilYeni.Domain
{
    /// <summary>
    /// Hidrolik band içi düşey kırık işareti (KotDeğişimÇizgisi).
    /// Class7’de 4m kırık => 40 CAD birimi (düşey ölçek 10 ise).
    /// </summary>
    public sealed class HydraulicVerticalBreakMarker
    {
        public int BandIndex { get; private set; }
        public double Km { get; private set; }
        public double FromTopRefMeters { get; private set; }
        public double ToTopRefMeters { get; private set; }

        public HydraulicVerticalBreakMarker(int bandIndex, double km, double fromTopRefMeters, double toTopRefMeters)
        {
            BandIndex = bandIndex;
            Km = km;
            FromTopRefMeters = fromTopRefMeters;
            ToTopRefMeters = toTopRefMeters;
        }
    }
}
