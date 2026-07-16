# Hoca Konuşmasına Göre Nihai Gereksinim Kontrol Listesi

| No | Hocanın isteği | Paketteki karşılığı | Durum |
|---:|---|---|:---:|
| 1 | Görev 1 / birim aktarma kullanılmasın | Kaynak `Ders Birim` / `Birim` doğrudan kullanılır | ✅ |
| 2 | Ders bazında analiz yapılsın | `Ders Kodu + Grup No` anahtarı | ✅ |
| 3 | 1. ve 2. sınıflar ayrılabilsin | Ders kodunun ilk rakamından sınıf | ✅ |
| 4 | Birinci, ikinci ve uzaktan öğretim ayrı görülsün | Kategori alanında öğretim türü | ✅ |
| 5 | 2024-2025 ve 2025-2026 yan yana olsun | Tüm karşılaştırma tablolarında çift dönem sütunları | ✅ |
| 6 | Katılımcı sayısı iki dönem ayrı gelsin | `Katılımcı 2024-2025 Güz` / `Katılımcı 2025-2026 Güz` | ✅ |
| 7 | Her birim/program ayrı Excel olsun | Memnuniyet: 19, başarı: 20 birim Excel’i | ✅ |
| 8 | Birimi boş olanlar ayrı dosya olsun | `Memnuniyet/Birimi_Bos.xlsx` | ✅ |
| 9 | Her Excel’in başında tüm dersler toplu görülsün | `01_Tum_Dersler_Ozet` | ✅ |
| 10 | Bölüm/kategori genel analizi olsun | `02_Bolum_Genel_Analiz` / `02_Birim_Genel_Analiz` | ✅ |
| 11 | Her ders ayrı ayrı görülsün | Memnuniyet dosyalarında `D_...` detay sayfaları | ✅ |
| 12 | Her dersin ortalaması, genel ortalama ve grup ortalaması | Ders detayları + bölüm genel analizi | ✅ |
| 13 | Görev 3’te MD Akademik tarzı ders grafiği | Her ders sayfasında iki dönemli sütun grafik | ✅ |
| 14 | Grafik başlığında ders bilgisi olsun | Ders kodu **ve ders adı** birlikte yazılır | ✅ |
| 15 | Grafikler 2024-2025 / 2025-2026 ikili olsun | Mavi ve turuncu yan yana sütunlar | ✅ |
| 16 | Word dışında üretilen Excel’ler teslim edilsin | `Sonuclar` altında 41 Excel | ✅ |
| 17 | Dosya yolu kolay ve taşınabilir olsun | `pathlib`, `Veriler`, CLI parametreleri, README | ✅ |
| 18 | Başkası da çalıştırabilsin | `calistir.bat`, `requirements.txt`, açıklamalar | ✅ |
| 19 | Başarıda ders kodu, grup ve ders adı | İki dönem karşılaştırmalı sütunlar | ✅ |
| 20 | Başarıda öğretim üyesi/hoca | İki dönem ayrı hoca sütunları | ✅ |
| 21 | Öğrenci, başarılı ve başarısız sayıları | İki dönem yan yana | ✅ |
| 22 | DZ/devamsız sayısı | İki dönem DZ sütunları | ✅ |
| 23 | Eski başarı oranı da kalsın | `Başarı Oranı` çift dönem sütunları | ✅ |
| 24 | Güncel başarı oranı eklensin | `Başarı Oranı Güncel` çift dönem sütunları | ✅ |
| 25 | Güncel oran DZ çıkarılarak hesaplansın | `Başarılı / (Öğrenci - DZ) × 100` | ✅ |
| 26 | Başarı çıktıları da birim bazında ayrı olsun | `Sonuclar/Basari_Oranlari` | ✅ |
| 27 | Başarı için karşılaştırmalı grafik olsun | Her başarı Excel’inde iki grafik sayfası/alanı | ✅ |
| 28 | Ağırlıklı ortalama açıklanabilsin | Hesaplama sayfası, README ve savunma notu | ✅ |
| 29 | Hesapların doğru olduğu kontrol edilsin | `00_Kontrol_Raporu.xlsx` | ✅ |
| 30 | Sonradan verilen güncel başarı sütunu doğrulansın | `01_Yeni_Dosya_Dogrulama.xlsx` | ✅ |
| 31 | Bahar da aynı kodla çalışabilsin | `--donem-1-adi`, `--donem-2-adi` ve dört dosya yolu | ✅ |

## Son teknik denetim

- Ana program temiz çıktı klasöründe çalıştırıldı: **çıkış kodu 0**.
- Memnuniyet birim Excel’i: **19**.
- Başarı birim Excel’i: **20**.
- Memnuniyet ders detay sayfası: **261**.
- Memnuniyet detay grafiği: **522** (her ders için 2 grafik).
- Başarı grafiği: **40** (her birim için 2 grafik).
- Yapısal denetimde eksik zorunlu sayfa/sütun/grafik: **0**.
- Resmî güncel oranla karşılaştırılan satır: **264**.
- Formül uyumsuzluğu: **0**.
- Paydası sıfır olduğu için ayrı raporlanan satır: **6**.
