# PlanProfilYeni – Proje Dokümantasyonu

## 1) Genel Bakış
Bu proje, Excel’deki hat verilerini okuyup hem Excel raporu üretir hem de AutoCAD üzerinde hidrolik/koordinat tabloları ve profil çizimleri üretir. Uygulama bir WinForms arayüzü üzerinden çalışır ve Excel ile AutoCAD’e COM üzerinden bağlanır.

- Uygulama girişi: [PlanProfilYeni/Program.cs](PlanProfilYeni/Program.cs)
- UI: [PlanProfilYeni/UI/Form1.cs](PlanProfilYeni/UI/Form1.cs)
- Proje yapılandırması: [PlanProfilYeni/PlanProfilYeni.csproj](PlanProfilYeni/PlanProfilYeni.csproj)

## 2) Mimari Katmanlar
### 2.1 UI (WinForms)
UI, Excel dosyalarını seçtirir ve 5 ana use‑case’i tetikler:
- Hidrolik tablo → Excel
- Hidrolik tablo → AutoCAD
- Koordinat tablosu → Excel
- Koordinat tablosu → AutoCAD
- Profil çizimi → AutoCAD

Ana form ve akışlar: [PlanProfilYeni/UI/Form1.cs](PlanProfilYeni/UI/Form1.cs)

### 2.2 Application (Use‑Case Katmanı)
Use‑case sınıfları UI’dan çağrılır ve servisleri/CAD yazıcılarını orkestre eder.
- `DrawProfileUseCase` → profil çizim akışının omurgası
- `ExportHydraulicExcelUseCase` → hidrolik Excel raporu
- `ExportCoordinateExcelUseCase` → koordinat Excel raporu
- `PrintHydraulicTableToCadUseCase` → hidrolik CAD tablosu
- `PrintCoordinateCadUseCase` → koordinat CAD tablosu
- `ImportExcelTablesUseCase` → hidrolik tablo import/export (alternatif akış)
- `DrawPlanUseCase` → plan çizimi için placeholder

Dosyalar: [PlanProfilYeni/Application/DrawProfileUseCase.cs](PlanProfilYeni/Application/DrawProfileUseCase.cs), [PlanProfilYeni/Application/ExportHydraulicExcelUseCase.cs](PlanProfilYeni/Application/ExportHydraulicExcelUseCase.cs), [PlanProfilYeni/Application/ExportCoordinateExcelUseCase.cs](PlanProfilYeni/Application/ExportCoordinateExcelUseCase.cs), [PlanProfilYeni/Application/PrintHydraulicTableToCadUseCase.cs](PlanProfilYeni/Application/PrintHydraulicTableToCadUseCase.cs), [PlanProfilYeni/Application/PrintCoordinateCadUseCase.cs](PlanProfilYeni/Application/PrintCoordinateCadUseCase.cs), [PlanProfilYeni/Application/ImportExcelTablesUseCase.cs](PlanProfilYeni/Application/ImportExcelTablesUseCase.cs), [PlanProfilYeni/Application/DrawPlanUseCase.cs](PlanProfilYeni/Application/DrawPlanUseCase.cs), [PlanProfilYeni/Application/HydraulicProcessOptions.cs](PlanProfilYeni/Application/HydraulicProcessOptions.cs)

### 2.3 Services (İş Mantığı + Excel Okuma/Yazma)
Excel’den veri okuma, profil/hidrolik yapılandırma ve rapor yazma burada yapılır.
- `HydraulicExcelService`: “Basınç Profili” sayfası satır/sütun okuma
- `HydraulicReportWriter`: hidrolik tabloyu yeni Excel’e yazma
- `CoordinateExcelService`: “Boru Plan Koordinatları” okuma
- `CoordinateReportWriter`: koordinat Excel çıktısı
- `ProfileExcelService`: “Arazi & Boru Profili” okuma
- `PressureProfileExcelService`: hidrolik seri ve ekipman listeleri
- `ProfileBandService`: band ve km kırıkları üretimi
- `ProfileClipService`: arazi/boru polylinelerini bandlara bölme
- `HydraulicProfileBuilder`: hidrolik segment ve dikey kırık üretimi
- `HydraulicVerticalBreakService`: eski kırık mantığı (destekleyici)
- `ProfileEquipmentBuilder`: ekipman listesi birleştirme
- `PipeExtremaDetector`: boru profili extrema (vantuz/tahliye)
- `GeometryService`: koordinat tablosu açı hesapları

Dosyalar: [PlanProfilYeni/Services/HydraulicExcelService.cs](PlanProfilYeni/Services/HydraulicExcelService.cs), [PlanProfilYeni/Services/HydraulicReportWriter.cs](PlanProfilYeni/Services/HydraulicReportWriter.cs), [PlanProfilYeni/Services/CoordinateExcelService.cs](PlanProfilYeni/Services/CoordinateExcelService.cs), [PlanProfilYeni/Services/CoordinateReportWriter.cs](PlanProfilYeni/Services/CoordinateReportWriter.cs), [PlanProfilYeni/Services/ProfileExcelService.cs](PlanProfilYeni/Services/ProfileExcelService.cs), [PlanProfilYeni/Services/PressureProfileExcelService.cs](PlanProfilYeni/Services/PressureProfileExcelService.cs), [PlanProfilYeni/Services/ProfileBandService.cs](PlanProfilYeni/Services/ProfileBandService.cs), [PlanProfilYeni/Services/ProfileClipService.cs](PlanProfilYeni/Services/ProfileClipService.cs), [PlanProfilYeni/Services/HydraulicProfileBuilder.cs](PlanProfilYeni/Services/HydraulicProfileBuilder.cs), [PlanProfilYeni/Services/HydraulicVerticalBreakService.cs](PlanProfilYeni/Services/HydraulicVerticalBreakService.cs), [PlanProfilYeni/Services/ProfileEquipmentBuilder.cs](PlanProfilYeni/Services/ProfileEquipmentBuilder.cs), [PlanProfilYeni/Services/PipeExtremaDetector.cs](PlanProfilYeni/Services/PipeExtremaDetector.cs), [PlanProfilYeni/Services/GeometryService.cs](PlanProfilYeni/Services/GeometryService.cs)

### 2.4 CAD Katmanı
AutoCAD model alanına tablo/profil çizimleri yapılır.
- `CadHydraulicTablePrinter` ve `CadHydraulicTableOptions`
- `CadCoordinateTablePrinter`
- `CadProfileGridPrinter`, `CadProfileHeaderPrinter`, `CadProfileFooterPrinter`
- `CadProfileTransformer`, `CadProfileTransformOptions`
- `CadProfilePolylinePrinter` (arazi/boru)
- `CadProfileHydraulicDrawer` (hidrolik eğri ve dikey kırık)
- `CadProfileEquipmentDrawer` (hidrant/BKV/vantuz/tahliye/hat ayrımı)

Dosyalar: [PlanProfilYeni/Cad/CadHydraulicTablePrinter.cs](PlanProfilYeni/Cad/CadHydraulicTablePrinter.cs), [PlanProfilYeni/Cad/CadHydraulicTableOptions.cs](PlanProfilYeni/Cad/CadHydraulicTableOptions.cs), [PlanProfilYeni/Cad/CadCoordinateTablePrinter.cs](PlanProfilYeni/Cad/CadCoordinateTablePrinter.cs), [PlanProfilYeni/Cad/CadProfileGridPrinter.cs](PlanProfilYeni/Cad/CadProfileGridPrinter.cs), [PlanProfilYeni/Cad/CadProfileHeaderPrinter.cs](PlanProfilYeni/Cad/CadProfileHeaderPrinter.cs), [PlanProfilYeni/Cad/CadProfileFooterPrinter.cs](PlanProfilYeni/Cad/CadProfileFooterPrinter.cs), [PlanProfilYeni/Cad/CadProfileTransformer.cs](PlanProfilYeni/Cad/CadProfileTransformer.cs), [PlanProfilYeni/Cad/CadProfileTransformOptions.cs](PlanProfilYeni/Cad/CadProfileTransformOptions.cs), [PlanProfilYeni/Cad/CadProfilePolylinePrinter.cs](PlanProfilYeni/Cad/CadProfilePolylinePrinter.cs), [PlanProfilYeni/Cad/CadProfileHydraulicDrawer.cs](PlanProfilYeni/Cad/CadProfileHydraulicDrawer.cs), [PlanProfilYeni/Cad/CadProfileEquipmentDrawer.cs](PlanProfilYeni/Cad/CadProfileEquipmentDrawer.cs), [PlanProfilYeni/Cad/ComRelease.cs](PlanProfilYeni/Cad/ComRelease.cs)

### 2.5 Domain (Model Katmanı)
Veri taşıyan sınıflar ve türler burada.
- `CoordinateRow`
- `HydraulicRow`
- `HydraulicBuildOptions`, `HydraulicValueMode`
- `HydraulicDrawableSegment`, `HydraulicVerticalBreakMarker`
- `ProfileBandSet`
- `ProfileEquipmentItem`, `ProfileEquipmentType`
- `ProfilePolylineBandSegment`

Dosyalar: [PlanProfilYeni/Domain/CoordinateRow.cs](PlanProfilYeni/Domain/CoordinateRow.cs), [PlanProfilYeni/Domain/HydraulicRow.cs](PlanProfilYeni/Domain/HydraulicRow.cs), [PlanProfilYeni/Domain/HydraulicBuildOptions.cs](PlanProfilYeni/Domain/HydraulicBuildOptions.cs), [PlanProfilYeni/Domain/HydraulicDrawableSegment.cs](PlanProfilYeni/Domain/HydraulicDrawableSegment.cs), [PlanProfilYeni/Domain/HydraulicVerticalBreakMarker.cs](PlanProfilYeni/Domain/HydraulicVerticalBreakMarker.cs), [PlanProfilYeni/Domain/ProfileBandSet.cs](PlanProfilYeni/Domain/ProfileBandSet.cs), [PlanProfilYeni/Domain/ProfileEquipmentItem.cs](PlanProfilYeni/Domain/ProfileEquipmentItem.cs), [PlanProfilYeni/Domain/ProfileEquipmentType.cs](PlanProfilYeni/Domain/ProfileEquipmentType.cs), [PlanProfilYeni/Domain/ProfilePolylineBandSegment.cs](PlanProfilYeni/Domain/ProfilePolylineBandSegment.cs)

## 3) Çalıştırma ve Ortam Gereksinimleri
- .NET Framework 4.8 (WinForms)
- AutoCAD 2026 Interop (COM)
- Microsoft Excel Interop (COM)

Referanslar ve COM bağımlılıkları: [PlanProfilYeni/PlanProfilYeni.csproj](PlanProfilYeni/PlanProfilYeni.csproj)

AutoCAD’in açık olması gerekir; COM ile çalışan uygulama aktif AutoCAD örneğini yakalar.

## 4) Kullanıcı Akışları (UI)
UI tek bir form üzerinden Excel seçimi ve çizim/rapor akışlarını tetikler.
- Excel ekleme/çıkarma listesi
- Hidrolik → Excel
- Hidrolik → AutoCAD
- Koordinat → Excel
- Koordinat → AutoCAD
- Profil → AutoCAD

Akışlar: [PlanProfilYeni/UI/Form1.cs](PlanProfilYeni/UI/Form1.cs)

## 5) Use‑Case Akışları
### 5.1 Hidrolik Tablo → Excel
1) UI `ExportHydraulicExcelUseCase` çağırır.
2) `HydraulicExcelService` Excel dosyalarını okur.
3) `HydraulicReportWriter` çıktı Excel oluşturur.

Dosyalar: [PlanProfilYeni/Application/ExportHydraulicExcelUseCase.cs](PlanProfilYeni/Application/ExportHydraulicExcelUseCase.cs), [PlanProfilYeni/Services/HydraulicExcelService.cs](PlanProfilYeni/Services/HydraulicExcelService.cs), [PlanProfilYeni/Services/HydraulicReportWriter.cs](PlanProfilYeni/Services/HydraulicReportWriter.cs)

### 5.2 Hidrolik Tablo → AutoCAD
1) UI `PrintHydraulicTableToCadUseCase` çağırır.
2) `HydraulicExcelService` okur, `CadHydraulicTablePrinter` çizdirir.
3) Grid ve yazı layer’ları `CadHydraulicTableOptions` ile yönetilir.

Dosyalar: [PlanProfilYeni/Application/PrintHydraulicTableToCadUseCase.cs](PlanProfilYeni/Application/PrintHydraulicTableToCadUseCase.cs), [PlanProfilYeni/Services/HydraulicExcelService.cs](PlanProfilYeni/Services/HydraulicExcelService.cs), [PlanProfilYeni/Cad/CadHydraulicTablePrinter.cs](PlanProfilYeni/Cad/CadHydraulicTablePrinter.cs), [PlanProfilYeni/Cad/CadHydraulicTableOptions.cs](PlanProfilYeni/Cad/CadHydraulicTableOptions.cs)

### 5.3 Koordinat Tablosu → Excel
1) UI `ExportCoordinateExcelUseCase` çağırır.
2) `CoordinateExcelService` okur, `CoordinateReportWriter` yazar.

Dosyalar: [PlanProfilYeni/Application/ExportCoordinateExcelUseCase.cs](PlanProfilYeni/Application/ExportCoordinateExcelUseCase.cs), [PlanProfilYeni/Services/CoordinateExcelService.cs](PlanProfilYeni/Services/CoordinateExcelService.cs), [PlanProfilYeni/Services/CoordinateReportWriter.cs](PlanProfilYeni/Services/CoordinateReportWriter.cs)

### 5.4 Koordinat Tablosu → AutoCAD
1) UI `PrintCoordinateCadUseCase` çağırır.
2) `CoordinateExcelService` okur, `CadCoordinateTablePrinter` çizdirir.

Dosyalar: [PlanProfilYeni/Application/PrintCoordinateCadUseCase.cs](PlanProfilYeni/Application/PrintCoordinateCadUseCase.cs), [PlanProfilYeni/Services/CoordinateExcelService.cs](PlanProfilYeni/Services/CoordinateExcelService.cs), [PlanProfilYeni/Cad/CadCoordinateTablePrinter.cs](PlanProfilYeni/Cad/CadCoordinateTablePrinter.cs)

### 5.5 Profil Çizimi → AutoCAD (En Kapsamlı Akış)
1) `DrawProfileUseCase` Excel’den arazi/boru profili ve basınç profili okur.
2) `ProfileBandService` band/break km üretir.
3) `CadProfileTransformer` + `CadProfileGridPrinter` ile grid çizimi yapılır.
4) `CadProfileHeaderPrinter` + `CadProfileFooterPrinter` band başlık/altlıkları basar.
5) `ProfileClipService` arazi/boru polylinelerini bandlara böler.
6) `CadProfilePolylinePrinter` arazi ve boru çizgilerini çizer.
7) `HydraulicProfileBuilder` hidrolik segment ve dikey kırıkları üretir.
8) `CadProfileHydraulicDrawer` hidrolik eğriyi ve kırıkları çizer.
9) `PressureProfileExcelService` ekipmanları okur.
10) `PipeExtremaDetector` boru profili extrema noktalarını çıkarır.
11) `ProfileEquipmentBuilder` ekipman listesini birleştirir.
12) `CadProfileEquipmentDrawer` ekipman sembollerini çizer.

Dosyalar: [PlanProfilYeni/Application/DrawProfileUseCase.cs](PlanProfilYeni/Application/DrawProfileUseCase.cs), [PlanProfilYeni/Services/ProfileExcelService.cs](PlanProfilYeni/Services/ProfileExcelService.cs), [PlanProfilYeni/Services/PressureProfileExcelService.cs](PlanProfilYeni/Services/PressureProfileExcelService.cs), [PlanProfilYeni/Services/ProfileBandService.cs](PlanProfilYeni/Services/ProfileBandService.cs), [PlanProfilYeni/Cad/CadProfileTransformer.cs](PlanProfilYeni/Cad/CadProfileTransformer.cs), [PlanProfilYeni/Cad/CadProfileGridPrinter.cs](PlanProfilYeni/Cad/CadProfileGridPrinter.cs), [PlanProfilYeni/Cad/CadProfileHeaderPrinter.cs](PlanProfilYeni/Cad/CadProfileHeaderPrinter.cs), [PlanProfilYeni/Cad/CadProfileFooterPrinter.cs](PlanProfilYeni/Cad/CadProfileFooterPrinter.cs), [PlanProfilYeni/Services/ProfileClipService.cs](PlanProfilYeni/Services/ProfileClipService.cs), [PlanProfilYeni/Cad/CadProfilePolylinePrinter.cs](PlanProfilYeni/Cad/CadProfilePolylinePrinter.cs), [PlanProfilYeni/Services/HydraulicProfileBuilder.cs](PlanProfilYeni/Services/HydraulicProfileBuilder.cs), [PlanProfilYeni/Cad/CadProfileHydraulicDrawer.cs](PlanProfilYeni/Cad/CadProfileHydraulicDrawer.cs), [PlanProfilYeni/Services/PipeExtremaDetector.cs](PlanProfilYeni/Services/PipeExtremaDetector.cs), [PlanProfilYeni/Services/ProfileEquipmentBuilder.cs](PlanProfilYeni/Services/ProfileEquipmentBuilder.cs), [PlanProfilYeni/Cad/CadProfileEquipmentDrawer.cs](PlanProfilYeni/Cad/CadProfileEquipmentDrawer.cs)

## 6) Excel Şablonları ve Kolon Beklentileri
### 6.1 Basınç Profili (Sheet: “Basınç Profili”)
- Hidrolik seri: başlangıç B4/C4, devam N/X sütunları (satır 8+)
- Hidrantlar: N (km), BH (çıkış sayısı)
- Hat ayrımı: N (km), BD içinde “ Ayr.”
- BKV: N (km), BD “BKV”, ilişkili değer I (r‑4)
- Manuel vantuz/tahliye: B12.. ve C12..

Kaynak: [PlanProfilYeni/Services/PressureProfileExcelService.cs](PlanProfilYeni/Services/PressureProfileExcelService.cs), [PlanProfilYeni/Services/HydraulicExcelService.cs](PlanProfilYeni/Services/HydraulicExcelService.cs)

### 6.2 Arazi & Boru Profili (Sheet: “Arazi & Boru Profili”)
- Arazi profili: B (km) / C (kot)
- Boru profili: D (km) / E (kot)

Kaynak: [PlanProfilYeni/Services/ProfileExcelService.cs](PlanProfilYeni/Services/ProfileExcelService.cs)

### 6.3 Boru Plan Koordinatları (Sheet: “Boru Plan Koordinatları”)
- Koordinatlar: H (X), I (Y)
- km artışı: E sütunu
- Dönüş/sapma: `GeometryService` ile hesaplanır

Kaynak: [PlanProfilYeni/Services/CoordinateExcelService.cs](PlanProfilYeni/Services/CoordinateExcelService.cs), [PlanProfilYeni/Services/GeometryService.cs](PlanProfilYeni/Services/GeometryService.cs)

## 7) CAD Çizim Detayları
### 7.1 Layer ve Blok İsimleri
- Hidrolik tablo grid layer: “Arazi”
- Hidrolik tablo text layer: “TipKesit-YazıÇizgileri”
- Koordinat tablo layer: “Arazi” ve “TipKesit-YazıÇizgileri”
- Profil grid layer: “Grid‑ince”, “Grid‑kalın” (veya tek `GridLayer`)
- Profil eğri layer: “TipKesit‑YazıÇizgileri”
- Dikey kırık layer: “KotDeğişimÇizgisi”

Ekipman blokları:
- “Hidrant” (dinamik blok, görünürlük “n ÇIKIŞ”)
- “BKV”
- “Vantuz”
- “Tahliye”

Kaynak: [PlanProfilYeni/Cad/CadHydraulicTablePrinter.cs](PlanProfilYeni/Cad/CadHydraulicTablePrinter.cs), [PlanProfilYeni/Cad/CadCoordinateTablePrinter.cs](PlanProfilYeni/Cad/CadCoordinateTablePrinter.cs), [PlanProfilYeni/Cad/CadProfileGridPrinter.cs](PlanProfilYeni/Cad/CadProfileGridPrinter.cs), [PlanProfilYeni/Cad/CadProfileHydraulicDrawer.cs](PlanProfilYeni/Cad/CadProfileHydraulicDrawer.cs), [PlanProfilYeni/Cad/CadProfileEquipmentDrawer.cs](PlanProfilYeni/Cad/CadProfileEquipmentDrawer.cs)

### 7.2 Ölçek ve Dönüşüm Varsayımları
- Bu projede “km” alanı pratikte metre kabul ediliyor.
- UI tarafında yatay ve düşey ölçek 1/5000 ve 1/100 değerlerinden, `cadUnitsPerMeter = 0.2` üretiliyor.
- Profil dönüşümleri `CadProfileTransformer` üzerinden `CadProfileTransformOptions` ile yönetiliyor.

Kaynak: [PlanProfilYeni/UI/Form1.cs](PlanProfilYeni/UI/Form1.cs), [PlanProfilYeni/Application/DrawProfileUseCase.cs](PlanProfilYeni/Application/DrawProfileUseCase.cs), [PlanProfilYeni/Cad/CadProfileTransformer.cs](PlanProfilYeni/Cad/CadProfileTransformer.cs), [PlanProfilYeni/Cad/CadProfileTransformOptions.cs](PlanProfilYeni/Cad/CadProfileTransformOptions.cs)

## 8) Önemli Teknik Notlar ve Varsayımlar
- AutoCAD açık değilse COM bağlantısı hata verir; UI bu durumu kullanıcıya bildirir.
- Excel sayfa adları ve kolon dizilimleri sabittir; farklı şablonlarla çalışmaz.
- Hidrolik dikey kırıklar 4 m adım üzerinden üretilir.
- `CadProfileHydraulicDrawer` içinde TopRef/Y‑offset konusu not edilmiş; mikro kırıkların Y ofseti proje gereksinimine göre netleştirilebilir.
- Excel/AutoCAD COM temizliği bazı sınıflarda manuel yapılır; açık Excel penceresi istenen akışlarda COM serbest bırakma bilinçli olarak ertelenmiştir.

Kaynak: [PlanProfilYeni/Application/PrintHydraulicTableToCadUseCase.cs](PlanProfilYeni/Application/PrintHydraulicTableToCadUseCase.cs), [PlanProfilYeni/Services/HydraulicProfileBuilder.cs](PlanProfilYeni/Services/HydraulicProfileBuilder.cs), [PlanProfilYeni/Cad/CadProfileHydraulicDrawer.cs](PlanProfilYeni/Cad/CadProfileHydraulicDrawer.cs), [PlanProfilYeni/Services/HydraulicReportWriter.cs](PlanProfilYeni/Services/HydraulicReportWriter.cs)

## 9) Veri Akışı Özeti
- Excel → Hidrolik tablo → Excel çıktısı
- Excel → Hidrolik tablo → AutoCAD çizimi
- Excel → Koordinat tablosu → Excel çıktısı
- Excel → Koordinat tablosu → AutoCAD çizimi
- Excel (Arazi/Boru + Basınç) → Profil bandları → CAD grid → Arazi/Boru polylineleri → Hidrolik eğri/kırık → Ekipman sembolleri

Bu akışların tamamı UI üzerinden tetiklenir ve use‑case katmanı tarafından yönetilir.

## 10) Sonraki Düzeltmeler İçin Referans Noktaları
Düzeltme planına geçerken özellikle şu alanlar kritik:
- Excel şablon kolon/sayfa bağımlılıkları
- Ölçek dönüşümü ve “km = metre” varsayımı
- AutoCAD layer ve blok adları
- Hidrolik dikey kırık Y‑offset davranışı
- COM yaşam döngüsü (AutoCAD/Excel açık kalma senaryoları)

Referans dosyalar: [PlanProfilYeni/UI/Form1.cs](PlanProfilYeni/UI/Form1.cs), [PlanProfilYeni/Application/DrawProfileUseCase.cs](PlanProfilYeni/Application/DrawProfileUseCase.cs), [PlanProfilYeni/Services/PressureProfileExcelService.cs](PlanProfilYeni/Services/PressureProfileExcelService.cs), [PlanProfilYeni/Cad/CadProfileHydraulicDrawer.cs](PlanProfilYeni/Cad/CadProfileHydraulicDrawer.cs)
