// Domain/ProfileBandSet.cs
using System;

namespace PlanProfilYeni.Domain
{
    /// <summary>
    /// Class7 smethod_7 band üretimi:
    /// - TopLevels  : gparam_3 (band üst kotları)
    /// - BaseLevels : gparam_4 (band alt kotları)
    /// - BreakKms   : gparam_5 (band geçiş km'leri) + en son 999999 sentinel
    /// </summary>
    public sealed class ProfileBandSet
    {
        public int[] TopLevels { get; private set; }
        public int[] BaseLevels { get; private set; }
        public double[] BreakKms { get; private set; }

        public int BandCount
        {
            get { return TopLevels != null ? TopLevels.Length : 0; }
        }

        public ProfileBandSet(int[] topLevels, int[] baseLevels, double[] breakKms)
        {
            if (topLevels == null) throw new ArgumentNullException(nameof(topLevels));
            if (baseLevels == null) throw new ArgumentNullException(nameof(baseLevels));
            if (breakKms == null) throw new ArgumentNullException(nameof(breakKms));
            if (topLevels.Length != baseLevels.Length)
                throw new ArgumentException("TopLevels and BaseLevels must have same length.");

            TopLevels = topLevels;
            BaseLevels = baseLevels;
            BreakKms = breakKms;
        }
    }
}
