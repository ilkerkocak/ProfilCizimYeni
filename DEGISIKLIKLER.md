# Profil Çizim Farkları ve Düzeltmeler

## 🎯 Yapılan Değişiklikler

### 1. Grid Ayarları
**Sorun:** Grid çizgileri çok seyrekti (10m aralıklar)
**Çözüm:** 
- İnce yatay grid: 1m aralıklarla
- Kalın yatay grid: 5m aralıklarla
- Dikey grid: 100m aralıklarla (değiştirilmedi)

```csharp
// DrawProfileUseCase.cs
GridThinLayer = "Grid-ince",
GridThickLayer = "Grid-kalin",
VerticalStepMeters = 100.0,        
HorizontalThinStepMeters = 1.0,    // ✅ 10.0 -> 1.0
HorizontalThickStepMeters = 5.0,   // ✅ Yeni eklendi
```

### 2. Layer Renkleri
**Sorun:** Grid çizgileri turuncu/sarı renkteydi
**Çözüm:** AutoCAD layer renkleri düzeltildi

```csharp
// DrawProfileUseCase.cs - Layer renk atamaları
EnsureLayer(doc, "Grid-ince", 5);        // Mavi
EnsureLayer(doc, "Grid-kalin", 5);       // Mavi
EnsureLayer(doc, "Arazi", 3);            // Yeşil
EnsureLayer(doc, "Boru", 1);             // Kırmızı
EnsureLayer(doc, "TipKesit-YazıÇizgileri", 7);  // Beyaz/Siyah
```

**AutoCAD Renk Kodları:**
- 1 = Kırmızı
- 3 = Yeşil
- 5 = Mavi
- 7 = Beyaz/Siyah (tema bazlı)

### 3. Ekipman Sembol Boyutları
**Sorun:** Semboller çok büyüktü (2.0 CAD birimi)
**Çözüm:** Sembol boyutu küçültüldü

```csharp
// CadProfileEquipmentDrawer.cs
SymbolSizeCad = 0.5;  // ✅ 2.0 -> 0.5
VerticalOffsetMeters = 2.0;  // ✅ 1.5 -> 2.0 (daha yukarıda)
```

### 4. Çizgi Kalınlıkları
**Sorun:** Profil çizgileri çok ince görünüyordu
**Çözüm:** Lineweight parametreleri eklendi

```csharp
// CadProfilePolylinePrinter.cs
ln.Lineweight = (AcadLineWeight)((int)(weight * 100));

// DrawProfileUseCase.cs
polyDrawer.DrawSegments(..., "Arazi", 0.25);  // İnce
polyDrawer.DrawSegments(..., "Boru", 0.5);    // Kalın
```

## 🔧 Hala Yapılması Gerekenler

### 1. Alt Tablo/Grid Yapısı
İlk resimde profilin altında bir tablo görünüyor (koordinat/hidrolik tablosu olabilir). Bu şu anda eksik.

**Gerekli işlem:**
- `CadCoordinateTablePrinter` veya `CadHydraulicTablePrinter` sınıflarını kullan
- DrawProfileUseCase içinde bu tabloları çağır

### 2. Ölçek ve Oranlar
Eğer çıktı hala farklı görünüyorsa:

```csharp
// DrawProfileUseCase.cs - Execute metodunda
horizontalScale = 0.2;  // 100m = 20 CAD birimi
verticalScale = 10.0;   // 1m = 10 CAD birimi (1/100 ölçek)
```

Bu değerleri istenen çıktıya göre ayarlayabilirsiniz.

### 3. Header ve Footer Formatı
Başlık ve alt bilgi formatlarının istenen çıktıya uygun olduğundan emin olun:
- `CadProfileHeaderPrinter.cs`
- `CadProfileFooterPrinter.cs`

## 📝 Test Etme

1. **Projeyi derleyin**
2. **AutoCAD'de çalıştırın**
3. **Kontrol edin:**
   - Grid çizgileri mavi mi?
   - Gridler yeterince sık mı? (1m ve 5m)
   - Ekipman sembolleri uygun boyutta mı?
   - Profil çizgileri (arazi-yeşil, boru-kırmızı) doğru renkte mi?

## 🎨 Layer Yapısı

| Layer Adı | Renk | Kullanım |
|-----------|------|----------|
| Grid-ince | Mavi (5) | İnce yatay gridler (1m) |
| Grid-kalin | Mavi (5) | Kalın gridler (5m ve dikey 100m) |
| Arazi | Yeşil (3) | Arazi profil çizgisi |
| Boru | Kırmızı (1) | Boru profil çizgisi |
| TipKesit-YazıÇizgileri | Beyaz/Siyah (7) | Ekipmanlar ve yazılar |

## 💡 İpuçları

- AutoCAD'de dark theme kullanıyorsanız, renklerin farklı görünebileceğini unutmayın
- LWDISPLAY komutuyla çizgi kalınlıklarını görünür yapın
- Layer Properties Manager'da renkleri kontrol edin
