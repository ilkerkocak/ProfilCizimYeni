// Domain/ProfileEquipmentType.cs
namespace PlanProfilYeni.Domain
{
    /// <summary>
    /// Class7 profil çıktısında kullanılan ekipman/işaret tipleri.
    /// </summary>
    public enum ProfileEquipmentType
    {
        Hydrant,        // BH dolu satır (km=N)
        Bkv,            // BD "BKV" (km=N)
        HatAyrimi,      // BD içinde " Ayr." (km=N, label=BD prefix)
        Vantuz,         // boru profilinden (maksimum)
        Tahliye         // boru profilinden (minimum)
    }
}
