namespace PlanProfilYeni.Domain;


public class HydraulicRow
{
    public string HatAdi { get; set; }
    public int KesitNo { get; set; }
    public string Kilometre { get; set; }
    public double BrutAlan { get; set; }
    public double SulamaModulu { get; set; }
    public double HatDebi { get; set; }
    public double Hiz { get; set; }
    public double HidrolikEğim { get; set; }
    public string BoruCinsi { get; set; }
    public double? IcCap { get; set; }
    public double? DisCap { get; set; }
    public double StatikBasinc { get; set; }
    public double DayanmaBasinci { get; set; }
    public string HidrantNo { get; set; }
    public double? HizmetAlan { get; set; }
    public double? Debi { get; set; }
    public int? CikisSayisi { get; set; }
    public string Tipi { get; set; }
    public double? DinamikBasinc { get; set; }

    public bool? BasincRegulatoru { get; set; }
    public bool? DebiLimitoru { get; set; }
}