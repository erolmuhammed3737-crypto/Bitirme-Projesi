# =============================================================================
# TBMYO Ders Anketi Analiz Scripti
# Hazırlayan: Selçuk Uzmanoğlu - 18.03.2026
#
# GÖREV 1: 2025-2026 Güz_basari_oranı.xlsx dosyasındaki Birim bilgisini
#          Ders Kodu + Grup No eşleşmesiyle tbmyo_2025-2026_guz.xlsx
#          dosyasının Ders Birim sütununa yazar.
#
# GÖREV 2: MD_Akademik.xlsx'teki Tum_Dersler_Ozet sayfası örnek alınarak
#          tbmyo verisinden her bölüm için ayrı özet sayfası oluşturur.
#
# KULLANIM:
#   1. pip install pandas openpyxl xlsxwriter
#   2. Bu .py dosyasını aşağıdaki xlsx dosyalarıyla AYNI KLASÖRE koyun:
#        - 2025-2026 Güz_basari_oranı.xlsx
#        - tbmyo_2025-2026_guz.xlsx
#   3. Terminalde:  python tbmyo_analiz.py
#   4. Çıktılar aynı klasörde oluşur:
#        - tbmyo_2025-2026_guz_updated.xlsx   (Ders Birim doldurulmuş)
#        - tbmyo_2025-2026_guz_Akademik.xlsx  (Özet sayfaları)
# =============================================================================

import pandas as pd
import re
import os

# --- Dosya Yolları ---
BASARI_FILE  = '2025-2026 Güz_basari_oranı.xlsx'
TBMYO_FILE   = 'tbmyo_2025-2026_guz.xlsx'
OUTPUT_TBMYO = 'tbmyo_2025-2026_guz_updated.xlsx'
OUTPUT_AKAD  = 'tbmyo_2025-2026_guz_Akademik.xlsx'

# --- Soru Grupları (MD_detay.py ile aynı) ---
QUESTION_GROUPS = {
    "Ders İçeriği":         ['1_1', '3_1', '14_1'],
    "Öğretim Elemanı":      ['2_1', '4_1', '5_1', '7_1', '9_1', '10_1', '11_1', '12_1', '14_1', '15_1', '16_1'],
    "Ölçme Değerlendirme":  ['6_1', '12_1', '13_1', '14_1'],
    "Yöntem":               ['4_1', '6_1', '8_1', '10_1', '14_1'],
}

SCOLS = ["Ders Birim", "Ders Kodu", "Kategori", "Katılımcı (N)"] + list(QUESTION_GROUPS.keys()) + ["Genel Memnuniyet"]

score_mapping = {
    'Kesinlikle katılmıyorum': 1, 'Katılmıyorum': 2,
    'Pek fazla katılmıyorum':  3, 'Biraz katılıyorum': 5,
    'Katılıyorum':             4, 'Tamamen katılıyorum': 6,
}
q6_map = {**score_mapping, 'Ödev, proje, ekip çalışması, öğrenci sunumları yapılmadı.': 0}
q8_map = {**score_mapping, 'Ders için kaynak önerilmedi.': 0}

# Renk paleti (bölüm satırları için)
ROW_PALETTE = [
    '#DAEEF3','#EBF1DE','#FDE9D9','#E4DFEC','#F2DCDB',
    '#FFFBCC','#DCE6F1','#F2F2F2','#D8E4BC','#FFC7CE',
    '#DDEBF7','#FCE4D6','#E2EFDA','#FFF2CC','#D6DCE4',
    '#EDEDED','#C6EFCE',
]

# =============================================================================
# YARDIMCI FONKSİYONLAR
# =============================================================================

def get_map(col):
    if col.startswith('6_1'): return q6_map
    if col.startswith('8_1'): return q8_map
    return score_mapping

def calc_groups(gdf):
    """Her soru grubu için ağırlıklı ortalama hesaplar."""
    res = {}
    for gname, prefixes in QUESTION_GROUPS.items():
        cols = [c for c in gdf.columns if any(c.startswith(p) for p in prefixes)]
        if not cols:
            continue
        t, n = 0, 0
        for q in cols:
            m = gdf[q].map(get_map(q))
            t += m.sum()
            n += m.count()
        res[gname] = round(t / n, 2) if n > 0 else 0
    return res

def calc_gen(gdf):
    """Tüm sorular için genel memnuniyet ortalaması."""
    cols = [c for c in gdf.columns if re.match(r'^\d+_1', c)]
    t, n = 0, 0
    for q in cols:
        m = gdf[q].map(get_map(q))
        t += m.sum()
        n += m.count()
    return round(t / n, 2) if n > 0 else 0

def identify_kat(ders_kodu, ders_birim):
    """Ders kodundan sınıf seviyesi ve öğretim tipini belirler."""
    m = re.search(r'[A-Za-z]+(\d)', str(ders_kodu))
    lvl = {"1": "1. Sınıf", "2": "2. Sınıf", "3": "3. Sınıf", "4": "4. Sınıf"}.get(
        m.group(1) if m else '', "Bilinmiyor"
    )
    tip = "Uzaktan Öğretim" if "Uzaktan Öğretim" in str(ders_birim) else "Birinci Öğretim"
    return f"{lvl} - {tip}"

def safe_sheet_name(name, maxlen=31):
    """Excel sayfa adı için geçersiz karakterleri temizler."""
    return re.sub(r'[\\/*?:\[\]]', '', name)[:maxlen]

def build_summary(df):
    """Ders Birim + Ders Kodu grubundan özet DataFrame üretir."""
    rows = []
    for (birim, kodu), gdf in df.groupby(['Ders Birim', 'Ders Kodu'], dropna=False):
        ga = calc_groups(gdf)
        if ga:
            rows.append({
                "Ders Birim":    birim,
                "Ders Kodu":     kodu,
                "Kategori":      identify_kat(kodu, birim),
                "Katılımcı (N)": len(gdf),
                **ga,
                "Genel Memnuniyet": calc_gen(gdf),
            })
    return pd.DataFrame(rows)[SCOLS] if rows else pd.DataFrame(columns=SCOLS)

def write_ozet_sheet(writer, ws_name, bdf, workbook, baslik):
    """Özet DataFrame'i formatlanmış şekilde Excel sayfasına yazar + grafik ekler."""
    bdf.to_excel(writer, sheet_name=ws_name, index=False, startrow=0)
    ws = writer.sheets[ws_name]

    hfmt = workbook.add_format({
        'bold': True, 'bg_color': '#D9D9D9', 'border': 1,
        'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
    })
    ws.set_row(0, 32)
    ws.set_column(0, 0, 38)   # Ders Birim
    ws.set_column(1, 1, 14)   # Ders Kodu
    ws.set_column(2, 2, 30)   # Kategori
    ws.set_column(3, 3, 14)   # Katılımcı (N)
    ws.set_column(4, 4, 16)   # Ders İçeriği
    ws.set_column(5, 5, 18)   # Öğretim Elemanı
    ws.set_column(6, 6, 22)   # Ölçme Değerlendirme
    ws.set_column(7, 7, 12)   # Yöntem
    ws.set_column(8, 8, 18)   # Genel Memnuniyet

    for col_idx, col_name in enumerate(bdf.columns):
        ws.write(0, col_idx, col_name, hfmt)

    birim_colors = {b: ROW_PALETTE[i % len(ROW_PALETTE)]
                    for i, b in enumerate(bdf['Ders Birim'].unique())}

    for row_idx, row in enumerate(bdf.itertuples(index=False), start=1):
        bg = birim_colors.get(row[0], '#FFFFFF')
        rfmt_str = workbook.add_format({'border': 1, 'align': 'left',   'valign': 'vcenter', 'bg_color': bg})
        rfmt_num = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': bg, 'num_format': '0.00'})
        rfmt_int = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': bg})
        for col_idx, val in enumerate(row):
            if col_idx < 3:
                ws.write(row_idx, col_idx, val, rfmt_str)
            elif col_idx == 3:
                ws.write(row_idx, col_idx, val, rfmt_int)
            else:
                ws.write(row_idx, col_idx, val if pd.notna(val) else '', rfmt_num)

    # Grafik
    if len(bdf) > 0:
        chart = workbook.add_chart({'type': 'column'})
        grup_cols = list(QUESTION_GROUPS.keys()) + ['Genel Memnuniyet']
        for gcol in grup_cols:
            ci = SCOLS.index(gcol)
            chart.add_series({
                'name':       gcol,
                'categories': [ws_name, 1, 1, len(bdf), 1],
                'values':     [ws_name, 1, ci, len(bdf), ci],
                'data_labels': {'value': True},
            })
        chart.set_title({'name': f'{baslik} – Memnuniyet Özeti'})
        chart.set_y_axis({'min': 0, 'max': 6, 'major_gridlines': {'visible': False}})
        chart.set_x_axis({'major_gridlines': {'visible': False}})
        chart.set_size({'x_scale': 2.5, 'y_scale': 1.5})
        ws.insert_chart(f'A{len(bdf) + 4}', chart)

# =============================================================================
# GÖREV 1 — Ders Birim Doldurma
# =============================================================================
print("=" * 60)
print("GÖREV 1: Ders Birim sütunu dolduruluyor...")
print("=" * 60)

if not os.path.exists(BASARI_FILE):
    print(f"HATA: '{BASARI_FILE}' bulunamadı!")
    exit(1)
if not os.path.exists(TBMYO_FILE):
    print(f"HATA: '{TBMYO_FILE}' bulunamadı!")
    exit(1)

basari = pd.read_excel(BASARI_FILE)
tbmyo  = pd.read_excel(TBMYO_FILE)

# Anahtar oluştur: Ders Kodu + Grup No
basari_clean = basari[['Ders Kodu', 'Grup No', 'Birim']].drop_duplicates()
basari_clean['Grup No'] = basari_clean['Grup No'].fillna(0).astype(int)
basari_clean['Ders Kodu'] = basari_clean['Ders Kodu'].astype(str).str.strip()
basari_lookup = (
    basari_clean
    .drop_duplicates(subset=['Ders Kodu', 'Grup No'])
    .set_index(['Ders Kodu', 'Grup No'])['Birim']
)

tbmyo['Ders Kodu'] = tbmyo['Ders Kodu'].astype(str).str.strip()
tbmyo['Grup No']   = tbmyo['Grup No'].astype(int)

onceki_null = tbmyo['Ders Birim'].isna().sum()

def fill_birim(row):
    if pd.isna(row['Ders Birim']) or str(row['Ders Birim']).strip() == '':
        return basari_lookup.get((row['Ders Kodu'], row['Grup No']), row['Ders Birim'])
    return row['Ders Birim']

tbmyo['Ders Birim'] = tbmyo.apply(fill_birim, axis=1)

sonraki_null = tbmyo['Ders Birim'].isna().sum()
print(f"  Önceki boş satır: {onceki_null}")
print(f"  Doldurulan satır: {onceki_null - sonraki_null}")
print(f"  Kalan boş satır : {sonraki_null}  (ATA121, TRD121, YDZİ121 gibi ortak dersler)")

tbmyo.to_excel(OUTPUT_TBMYO, index=False)
print(f"\n  ✓ Kaydedildi: {OUTPUT_TBMYO}")

# =============================================================================
# GÖREV 2 — Bölüm Bazlı Özet Sayfaları
# =============================================================================
print()
print("=" * 60)
print("GÖREV 2: Bölüm özet sayfaları oluşturuluyor...")
print("=" * 60)

df = tbmyo.copy()
df.columns = df.columns.str.strip().str.replace('\u200b', '', regex=False).str.replace('\xa0', '', regex=False)
df['Ders Birim'] = df['Ders Birim'].fillna('').astype(str).str.strip()
df = df[df['Ders Birim'] != '']

birimler = sorted(df['Ders Birim'].unique())
print(f"  Toplam geçerli satır : {len(df)}")
print(f"  Bölüm (Ders Birim) sayısı: {len(birimler)}")

# Tüm bölümleri kapsayan özet
summary_df = build_summary(df)

# Her birim için ayrı özet
birim_summaries = {}
for birim in birimler:
    bdf = build_summary(df[df['Ders Birim'] == birim])
    if len(bdf) > 0:
        birim_summaries[birim] = bdf

with pd.ExcelWriter(OUTPUT_AKAD, engine='xlsxwriter') as writer:
    workbook = writer.book

    # 1. Tum_Dersler_Ozet
    write_ozet_sheet(writer, 'Tum_Dersler_Ozet', summary_df, workbook, 'Tüm Bölümler')
    print(f"  ✓ Sayfa: Tum_Dersler_Ozet  ({len(summary_df)} ders)")

    # 2. Her bölüm için ayrı sayfa
    used_names = {'Tum_Dersler_Ozet'}
    for birim, bdf in birim_summaries.items():
        sname = safe_sheet_name(birim)
        base, i = sname, 1
        while sname in used_names:
            sname = f"{base[:28]}_{i}"
            i += 1
        used_names.add(sname)
        write_ozet_sheet(writer, sname, bdf, workbook, birim)
        print(f"  ✓ Sayfa: {sname:<35} ({len(bdf)} ders)")

print(f"\n  ✓ Kaydedildi: {OUTPUT_AKAD}")
print()
print("=" * 60)
print("TÜM İŞLEMLER TAMAMLANDI!")
print("=" * 60)
