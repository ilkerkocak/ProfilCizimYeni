// Domain/ProfileEquipmentItem.cs
using System;

namespace PlanProfilYeni.Domain
{
    /// <summary>
    /// Profil ekipmanı/işareti.
    /// Bazı tipler ek bilgi taşır:
    /// - Hydrant: OutletCount (BH)
    /// - HatAyrimi: Label (BD prefix)
    /// - Bkv: StaticWaterLevel (I ile ilişkilendirilen değer) opsiyonel
    /// </summary>
    public sealed class ProfileEquipmentItem
    {
        public ProfileEquipmentType Type { get; private set; }
        public double Km { get; private set; }

        public int? OutletCount { get; private set; }          // Hydrant (BH)
        public string Label { get; private set; }              // HatAyrimi
        public double? StaticWaterLevel { get; private set; }  // Bkv ile ilişkilendirilen SSK

        public ProfileEquipmentItem(ProfileEquipmentType type, double km, int? outletCount, string label, double? staticWaterLevel)
        {
            Type = type;
            Km = km;
            OutletCount = outletCount;
            Label = label;
            StaticWaterLevel = staticWaterLevel;
        }

        public static ProfileEquipmentItem Hydrant(double km, int outletCount)
            => new ProfileEquipmentItem(ProfileEquipmentType.Hydrant, km, outletCount, null, null);

        public static ProfileEquipmentItem Bkv(double km, double? staticWaterLevel)
            => new ProfileEquipmentItem(ProfileEquipmentType.Bkv, km, null, null, staticWaterLevel);

        public static ProfileEquipmentItem HatAyrimi(double km, string label)
            => new ProfileEquipmentItem(ProfileEquipmentType.HatAyrimi, km, null, label, null);

        public static ProfileEquipmentItem Vantuz(double km)
            => new ProfileEquipmentItem(ProfileEquipmentType.Vantuz, km, null, null, null);

        public static ProfileEquipmentItem Tahliye(double km)
            => new ProfileEquipmentItem(ProfileEquipmentType.Tahliye, km, null, null, null);
    }
}
