# TBMYO İki Dönem Karşılaştırmalı Akademik Analiz

Bu proje, **2024-2025 Güz** ve **2025-2026 Güz** dönemlerine ait ders memnuniyet anketleri ile başarı oranı raporlarını karşılaştırır. Kaynak dosyalardaki `Ders Birim / Birim` alanına göre her program için ayrı Excel dosyaları oluşturur.

## Hocanın istediği düzeltmeler

- Eski **Görev 1 – Birim Aktarma** kullanılmaz.
- Kaynak dosyadaki birim bilgisi doğrudan kullanılır.
- Birimi boş olan anket satırları `Birimi_Bos.xlsx` dosyasında ayrıca raporlanır.
- İki dönem aynı tabloda, yan yana gösterilir.
- Her birim için ayrı Excel oluşturulur.
- Her memnuniyet Excel’inin başında:
  1. `01_Tum_Dersler_Ozet`
  2. `02_Bolum_Genel_Analiz`
  sayfaları bulunur.
- Her ders kodu ve grup numarası için ayrı detay sayfası ve karşılaştırmalı grafik oluşturulur. Grafik başlığında **ders kodu ve ders adı** birlikte yazılır.
- Katılımcı sayıları iki dönem için ayrı gösterilir.
- Dersler sınıf ve öğretim türüne göre gruplandırılır.
- Başarı analizinde öğrenci, başarılı, başarısız, DZ, kaynak başarı oranı ve güncel başarı oranı yer alır.
- Güncel başarı oranı devamsız öğrenciler paydadan çıkarılarak hesaplanır.
- Hocanın gönderdiği `2024-2025_Güz_güncel.xlsx` dosyasındaki `Başarı oranı_Güncel (%)` sütunu satır bazında yeniden hesaplanarak doğrulanır.

## Tek tıkla çalıştırma

Önce ZIP dosyasını **Tümünü Ayıkla** seçeneğiyle normal bir klasöre çıkarın. ZIP’in içinden doğrudan çalıştırmayın.

Windows’ta proje klasöründeki **`calistir.bat`** dosyasına çift tıklayın.

Program:

1. Önce `py -3`, ardından `python` komutuyla Python kurulumunu kontrol eder.
2. Python yoksa resmi indirme sayfasını açıp kurulumu nasıl yapacağınızı ekranda gösterir.
3. Proje klasöründe `.venv` adlı izole çalışma ortamı oluşturur.
4. Gerekli Python kütüphanelerini bu ortama kurar.
5. `Veriler` klasöründeki dört Excel’i otomatik bulur.
6. Analizi çalıştırır ve `Sonuclar` klasörünü açar.

İlk çalıştırmada Python paketlerinin indirilebilmesi için internet bağlantısı gerekir. Sonraki çalıştırmalarda aynı `.venv` kullanılır.

Kodun içine `C:\...` biçiminde sabit bilgisayar yolu yazılmamıştır. Dosya yolları `pathlib.Path` ile proje klasörüne göre otomatik çözülür. Bu nedenle proje başka bilgisayara kopyalandığında kod değişikliği gerekmez.

## Komut satırından çalıştırma

```bash
python -m pip install -r requirements.txt
python bitirme_projesi_analiz.py
```

Farklı dosyalarla çalıştırma:

```bash
python bitirme_projesi_analiz.py ^
  --memnuniyet-1 "C:\Veriler\2024-2025_guz_anket.xlsx" ^
  --memnuniyet-2 "C:\Veriler\2025-2026_guz_anket.xlsx" ^
  --basari-1 "C:\Veriler\2024-2025_guz_basari.xlsx" ^
  --basari-2 "C:\Veriler\2025-2026_guz_basari.xlsx" ^
  --donem-1-adi "2024-2025 Güz" ^
  --donem-2-adi "2025-2026 Güz" ^
  --cikti "C:\Veriler\Sonuclar"
```

Komut seçenekleri:

```bash
python bitirme_projesi_analiz.py --help
```

### Bahar döneminde aynı kodu kullanma

Kod içindeki dönem sabitlerini değiştirmek gerekmez. Dört Bahar dosyasının yolunu ve dönem adlarını komut satırından verin:

```bash
python bitirme_projesi_analiz.py ^
  --memnuniyet-1 "C:\Veriler\2024-2025_bahar_anket.xlsx" ^
  --memnuniyet-2 "C:\Veriler\2025-2026_bahar_anket.xlsx" ^
  --basari-1 "C:\Veriler\2024-2025_bahar_basari.xlsx" ^
  --basari-2 "C:\Veriler\2025-2026_bahar_basari.xlsx" ^
  --donem-1-adi "2024-2025 Bahar" ^
  --donem-2-adi "2025-2026 Bahar" ^
  --cikti "C:\Veriler\Bahar_Sonuclari"
```

Bu yapı, hocanın “Bahar dönemini de aynı kodla çalıştırma” beklentisini kod düzenlemeden karşılar.

## Girdi dosyaları

`Veriler` klasöründe şu sade dosya adları kullanılmıştır:

- `memnuniyet_2024_2025_guz.xlsx`
- `memnuniyet_2025_2026_guz.xlsx`
- `basari_2024_2025_guz.xlsx`
- `basari_2025_2026_guz.xlsx`

Yeni dönemlerde aynı dosya adlarının içeriği değiştirilebilir veya komut satırından farklı yollar verilebilir.

### Memnuniyet dosyasında gerekli alanlar

Memnuniyet kaynaklarında `Ders Adı` alanı bulunmadığından program, aynı dönemin başarı dosyasındaki `Ders Kodu + Grup No` eşleşmesinden ders adını alır. Eşleşme yoksa ders kodu korunur ve veri uydurulmaz.

- `Ders Üst Birim`
- `Ders Birim`
- `Ders Alt Birim`
- `Ders Kodu`
- `Grup No`
- `Öğretim Üyesi`
- `1_1` ile `16_1` arasında başlayan soru sütunları

### Başarı dosyasında gerekli alanlar

- `Üst Birim`
- `Birim`
- `Ders Kodu`
- `Grup No`
- `Ders Adı`
- `Öğretim Üyesi`
- `Öğrenci Sayısı`
- `Başarılı Öğrenci Sayısı`
- `Başarısız Öğrenci Sayısı`
- `Başarı Oranı (%)` veya `Başarı Oranı(%)`
- `DZ`

Başarı dosyasındaki tüm üniversite verileri içinden yalnızca **Teknik Bilimler Meslek Yüksekokulu** otomatik süzülür.

2024-2025 Güz için hocanın gönderdiği güncel dosya, teslim paketinde taşınabilir adla `basari_2024_2025_guz.xlsx` olarak saklanır. Dosyadaki ortak 32 sütun eski kaynakla birebir aynıdır; ek olarak resmi `Başarı oranı_Güncel (%)` sütunu bulunur.

## Çıktı yapısı

```text
Sonuclar/
├── 00_Kontrol_Raporu.xlsx
├── 01_Yeni_Dosya_Dogrulama.xlsx
├── Memnuniyet/
│   ├── Bilgisayar_Programciligi.xlsx
│   ├── Elektrik.xlsx
│   ├── Moda_Tasarimi.xlsx
│   ├── Birimi_Bos.xlsx
│   └── ...
└── Basari_Oranlari/
    ├── Bilgisayar_Programciligi.xlsx
    ├── Elektrik.xlsx
    ├── Moda_Tasarimi.xlsx
    └── ...
```

### Memnuniyet birim Excel’i

- `01_Tum_Dersler_Ozet`: Her dersin iki dönem katılımcı sayıları, grup ortalamaları, genel memnuniyet ve farkları.
- `02_Bolum_Genel_Analiz`: 1. sınıf, 2. sınıf, birinci/ikinci/uzaktan öğretim ve bölüm geneli karşılaştırması.
- `03_Hesaplama_Aciklama`: Kullanılan yöntem ve ortalama açıklamaları.
- `D_...` sayfaları: Her ders için ders kodu + ders adı başlığı, kategori tablosu, soru bazlı tablo ve iki karşılaştırmalı grafik.

### Başarı birim Excel’i

- `01_Tum_Dersler_Ozet`: Ders kodu, grup, ders adı, öğretim üyesi ve iki dönemin bütün başarı alanları.
- `02_Birim_Genel_Analiz`: Kategori ve birim toplamları.
- `03_Formul_Aciklama`: Formüllerin açık tanımı.
- `04_Grafikler`: Ders bazında iki dönem güncel başarı oranı grafiği.

`00_Kontrol_Raporu.xlsx` yeniden çalıştırıldığında `04_Guncel_Oran_Kontrol` sayfasında resmi güncel oran sütununun formülle uyumu da gösterilir. Pakette ayrıca ayrıntılı `01_Yeni_Dosya_Dogrulama.xlsx` bulunur.

## Memnuniyet hesaplaması

Likert puanları:

| Cevap | Puan |
|---|---:|
| Kesinlikle katılmıyorum | 1 |
| Katılmıyorum | 2 |
| Pek fazla katılmıyorum | 3 |
| Katılıyorum | 4 |
| Biraz katılıyorum | 5 |
| Tamamen katılıyorum | 6 |

Referans `MD_detay.py` koduna uygun özel cevaplar:

- `6_1`: “Ödev, proje, ekip çalışması, öğrenci sunumları yapılmadı.” → **0**
- `8_1`: “Ders için kaynak önerilmedi.” → **0**

Soru grupları:

- **Ders İçeriği:** `1_1, 3_1, 14_1`
- **Öğretim Elemanı:** `2_1, 4_1, 5_1, 7_1, 9_1, 10_1, 11_1, 12_1, 14_1, 15_1, 16_1`
- **Ölçme Değerlendirme:** `6_1, 12_1, 13_1, 14_1`
- **Yöntem:** `4_1, 6_1, 8_1, 10_1, 14_1`
- **Genel Memnuniyet:** `1_1–16_1` arasındaki bütün geçerli cevaplar

## “Ağırlıklı ortalama” ne demektir?

Bölüm genelinde ders ortalamalarının basit ortalaması alınmaz. Bütün geçerli öğrenci cevaplarının toplamı, geçerli cevap sayısına bölünür:

```text
Bölüm Ortalaması = Geçerli Puanların Toplamı / Geçerli Cevap Sayısı
```

Böylece 100 katılımcılı ders ile 5 katılımcılı ders bölümü eşit ölçüde etkilemez. Her öğrenci cevabı eşit ağırlıktadır; katılımcısı çok olan ders doğal olarak bölüm sonucuna daha fazla katkı verir.

## Başarı oranı formülü

Kaynak dosyadaki oran ayrıca korunur. Güncel oran şu formülle yeniden hesaplanır:

```text
Devamsız Hariç Öğrenci Sayısı = Öğrenci Sayısı - DZ

Başarı Oranı Güncel =
Başarılı Öğrenci Sayısı / Devamsız Hariç Öğrenci Sayısı × 100
```

Payda sıfırsa oran boş bırakılır; program sıfıra bölme hatası vermez.

Resmi güncel oran sütunu mevcutsa program bu değeri aynı formülle yeniden hesaplar. Sayısal olarak karşılaştırılabilen satırlarda `0,01` puandan büyük fark varsa terminalde uyarı verir. Paydası sıfır olan satırlar doğrulama raporunda ayrıca gösterilir.

## Derslerin eşleştirilmesi

İki dönem şu ortak anahtarla eşleştirilir:

```text
Ders Kodu + Grup No
```

Bir ders yalnızca bir dönemde bulunuyorsa diğer dönemin hücreleri boş bırakılır. Veri uydurulmaz.

## Sınıf ve öğretim türü

- Ders kodunda harflerden sonra gelen ilk rakam sınıfı belirler.
- Birim/alt birim metninde `Uzaktan` varsa **Uzaktan Öğretim**,
- `(İÖ)` varsa **İkinci Öğretim**,
- diğerleri **Birinci Öğretim** olarak sınıflandırılır.

## Sorun giderme

### `python` bulunamadı

Python 3.10 veya üzerini kurup kurulum ekranında **Add Python to PATH** seçeneğini işaretleyin.

### Excel dosyası açık hatası

Çıktı dosyalarını Excel’de kapatıp programı yeniden çalıştırın.

### Sütun eksik hatası

Hata mesajında yazan sütun adını kaynak Excel’in ilk satırında kontrol edin. Kod boşluk ve görünmez karakterleri temizler; ancak alanın kendisi bulunmalıdır.

### Eski çıktıları silmek

`temiz_sonuclar.bat` dosyasına çift tıklayın.
