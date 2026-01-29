// Domain/HydraulicBuildOptions.cs
using System;

namespace PlanProfilYeni.Domain
{
    public enum HydraulicValueMode
    {
        /// <summary>Değer zaten "kot" (m) şeklinde; direkt çiz.</summary>
        AbsoluteElevationMeters = 0,

        /// <summary>Değer "basınç yüksekliği" (m) gibi; PipeElevation + Value çiz.</summary>
        AddToPipeElevationMeters = 1
    }

    public sealed class HydraulicBuildOptions
    {
        /// <summary>Class7: topLimit her 4.0 birimde bir düşürülür.</summary>
        public double VerticalBreakStepMeters { get; set; } = 4.0;

        /// <summary>Hidrolik değerin nasıl yorumlanacağı.</summary>
        public HydraulicValueMode ValueMode { get; set; } = HydraulicValueMode.AbsoluteElevationMeters;

        /// <summary>
        /// Eğer ValueMode=AddToPipeElevationMeters ise, Value'nin birimi m değilse (örn bar) dönüşüm faktörü.
        /// 1 bar ≈ 10.197 mSS; projene göre netleştirince set edersin.
        /// </summary>
        public double ValueToMetersFactor { get; set; } = 1.0;

        public void Validate()
        {
            if (VerticalBreakStepMeters <= 0) throw new ArgumentOutOfRangeException(nameof(VerticalBreakStepMeters));
            if (ValueToMetersFactor <= 0) throw new ArgumentOutOfRangeException(nameof(ValueToMetersFactor));
        }
    }
}
