import pandas as pd
import os
import re

# --- YAPI VE DOSYA AYARLARI ---
# Orijinal veri dosyası ve çıktı dosyası
INPUT_EXCEL_FILE = 'MD.xlsx'  # Orijinal analiz veri seti
OUTPUT_FILE = 'MD_Akademik.xlsx'  # Analiz sonuçlarının raporu

# Yeni veri dosyası (hocanın gönderdiği)
TBMYO_INPUT_FILE = 'tbmyo_2025-2026_guz.xlsx'

# Gruplama için kullanılacak kolonlar
GROUPING_COLUMNS = ["Ders Birim", "Ders Kodu"]

# Soru grupları ve renklendirme
QUESTION_GROUPS = {
    "Ders İçeriği": ['1_1', '3_1', '14_1'],
    "Öğretim Elemanı": ['2_1', '4_1', '5_1', '7_1', '9_1', '10_1', '11_1', '12_1', '14_1','15_1', '16_1'],
    "Ölçme Değerlendirme": ['6_1', '12_1', '13_1','14_1'],
    "Yöntem": ['4_1', '6_1', '8_1', '10_1', '14_1']
}

GROUP_COLORS = {
    "Ders İçeriği": "#E6B8B7", "Öğretim Elemanı": "#B7DEE8",
    "Ölçme Değerlendirme": "#CCC0DA", "Yöntem": "#FCD5B5",
    "Memnuniyet Ortalaması": "#92D050"
}

# --- PUANLAMA SÖZLÜKLERİ ---
score_mapping = {'Kesinlikle katılmıyorum': 1, 'Katılmıyorum': 2, 'Pek fazla katılmıyorum': 3, 'Biraz katılıyorum': 5, 'Katılıyorum': 4, 'Tamamen katılıyorum': 6}
q6_score_mapping = {**score_mapping, 'Ödev, proje, ekip çalışması, öğrenci sunumları yapılmadı.': 0}
q8_score_mapping = {**score_mapping, 'Ders için kaynak önerilmedi.': 0}

# --- GRAFİK FONKSİYONU ---
def create_chart_original_format(workbook, sheet_name, categories_range, values_range, title, is_line=False):
    chart_type = 'line' if is_line else 'column'
    chart = workbook.add_chart({'type': chart_type})
    series_params = {
        'categories': f"='{sheet_name}'!{categories_range}",
        'values':     f"='{sheet_name}'!{values_range}",
        'data_labels': {'value': True},
    }
    if not is_line:
        series_params['fill'] = {'color': '#4F81BD'}
    else:
        series_params['line'] = {'color': '#ED7D31', 'width': 2.25}
        series_params['marker'] = {'type': 'circle', 'size': 6}
    chart.add_series(series_params)
    chart.set_title({'name': title})
    chart.set_x_axis({'major_gridlines': {'visible': False}})
    chart.set_y_axis({'visible': True, 'major_gridlines': {'visible': False}, 'min': 0, 'max': 6})
    chart.set_legend({'none': True})
    return chart

# --- SINIF VE ÖĞRETİM TÜRÜ BELİRLEME ---
def identify_class_and_type(ders_kodu, ders_birim):
    class_level = "Bilinmiyor"
    match = re.search(r'[A-Za-z]+(\d)', str(ders_kodu))
    if match:
        digit = match.group(1)
        if digit == '1': class_level = "1. Sınıf"
        elif digit == '2': class_level = "2. Sınıf"
        elif digit == '3': class_level = "3. Sınıf"
        elif digit == '4': class_level = "4. Sınıf"
    teaching_type = "Uzaktan Öğretim" if "Uzaktan Öğretim" in str(ders_birim) else "Birinci Öğretim"
    return f"{class_level} - {teaching_type}"

# --- GRUP ORTALAMALARI HESAPLAMA ---
def calculate_group_averages(group_df):
    group_results = {}
    for group_name, prefixes in QUESTION_GROUPS.items():
        group_questions = [col for col in group_df.columns if any(col.startswith(prefix) for prefix in prefixes)]
        if not group_questions: continue
        total, count = 0, 0
        for q in group_questions:
            mapping = q6_score_mapping if q.startswith('6_1') else (q8_score_mapping if q.startswith('8_1') else score_mapping)
            mapped = group_df[q].map(mapping)
            total += mapped.sum(); count += mapped.count()
        group_results[group_name] = round(total / count, 2) if count > 0 else 0
    return group_results

def calculate_question_averages(group_df):
    question_results = {}
    relevant_cols = [col for col in group_df.columns if "_1" in col and col[0].isdigit()]
    relevant_cols.sort(key=lambda x: int(x.split('_')[0]))
    for q in relevant_cols:
        mapping = q6_score_mapping if q.startswith('6_1') else (q8_score_mapping if q.startswith('8_1') else score_mapping)
        mapped = group_df[q].map(mapping)
        avg = mapped.mean()
        if not pd.isna(avg): question_results[q] = round(avg, 2)
    return question_results

def calculate_generic_avg(group_df):
    memnuniyet_sorulari = [col for col in group_df.columns if any(col.startswith(f"{i}_1") for i in range(1, 17))]
    total, count = 0, 0
    for q in memnuniyet_sorulari:
        mapping = q6_score_mapping if q.startswith('6_1') else (q8_score_mapping if q.startswith('8_1') else score_mapping)
        mapped = group_df[q].map(mapping)
        total += mapped.sum(); count += mapped.count()
    return round(total / count, 2) if count > 0 else 0

# --- YARIYIL BELİRLEME ---
def get_semester(ders_kodu):
    if len(str(ders_kodu))>=4:
        if str(ders_kodu)[3]=='1': return "1. Yarıyıl"
        if str(ders_kodu)[3]=='2': return "3. Yarıyıl"
    return "Bilinmiyor"

# --- ANA İŞLEM ---
print("Analiz Başlatılıyor...")

try:
    df = pd.read_excel(INPUT_EXCEL_FILE)
    df.columns = df.columns.str.strip().str.replace('\u200b', '', regex=False).str.replace('\xa0', '', regex=False)
    for col in GROUPING_COLUMNS:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.replace('  ', ' ')
    df['Kategori'] = df.apply(lambda row: identify_class_and_type(row['Ders Kodu'], row['Ders Birim']), axis=1)
except Exception as e:
    print(f"Hata: {e}"); exit(1)

summary_list = []

# --- HOCANIN İSTEDİĞİ TBMYO VERİSİ İŞLEME (PROGRAM BAZLI DOSYA) ---
try:
    tbmyo_df = pd.read_excel(TBMYO_INPUT_FILE)
    tbmyo_df.columns = tbmyo_df.columns.str.strip()
    for program in tbmyo_df['Ders Birim'].unique():
        prog_df = tbmyo_df[tbmyo_df['Ders Birim']==program]
        safe_name = "".join([c if c.isalnum() else "_" for c in str(program)])
        output_file_name = f"{safe_name}_Analiz_Raporu.xlsx"
        with pd.ExcelWriter(output_file_name, engine='xlsxwriter') as writer:
            workbook = writer.book
            prog_df['Yariyil'] = prog_df['Ders Kodu'].apply(get_semester)
            for yarıyıl in ["1. Yarıyıl","3. Yarıyıl"]:
                sem_df = prog_df[prog_df['Yariyil']==yarıyıl]
                if sem_df.empty: continue
                # Grup No bazlı hesaplama
                analiz_list = []
                for (ders_kodu, grup_no), g_df in sem_df.groupby(['Ders Kodu','Grup No']):
                    gen_avg = calculate_generic_avg(g_df)
                    analiz_list.append({"Ders Kodu":ders_kodu,"Grup No":grup_no,"Genel Memnuniyet":gen_avg})
                final_df = pd.DataFrame(analiz_list)
                final_df.to_excel(writer, sheet_name=yarıyıl,index=False)
                ws = writer.sheets[yarıyıl]
                ws.set_column('A:C',20)
                chart = create_chart_original_format(workbook, yarıyıl, f"$A$2:$A${len(final_df)+1}", f"$C$2:$C${len(final_df)+1}", f"{program} - {yarıyıl} Analizi")
                ws.insert_chart('E2', chart)
except Exception as e:
    print(f"TBMYO verisi işlenirken hata: {e}")

# --- ORİJİNAL MD ANALİZ VE RAPOR ---
with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
    workbook = writer.book
    # 1. GENEL ANALİZ SAYFASI
    ozet_data=[]
    kategoriler = sorted(df['Kategori'].unique())
    for kat in kategoriler:
        kat_df = df[df['Kategori']==kat]
        gruplar = calculate_group_averages(kat_df)
        gruplar['Genel Memnuniyet'] = calculate_generic_avg(kat_df)
        gruplar['Kategori'] = kat
        ozet_data.append(gruplar)
    ozet_df = pd.DataFrame(ozet_data)
    cols_order = ['Kategori'] + list(QUESTION_GROUPS.keys()) + ['Genel Memnuniyet']
    ozet_df = ozet_df[cols_order]
    ozet_df.to_excel(writer, sheet_name='Bolum_Genel_Analiz', index=False)
    ws_ozet = writer.sheets['Bolum_Genel_Analiz']
    ws_ozet.set_column('A:F',20)
    toplu_chart = workbook.add_chart({'type':'column'})
    for i,col_name in enumerate(cols_order[1:],start=1):
        toplu_chart.add_series({
            'name':['Bolum_Genel_Analiz',0,i],
            'categories':['Bolum_Genel_Analiz',1,0,len(ozet_df),0],
            'values':['Bolum_Genel_Analiz',1,i,len(ozet_df),i],
            'data_labels':{'value':True}
        })
    toplu_chart.set_title({'name':'Bölüm Genel Memnuniyet Kıyaslaması'})
    toplu_chart.set_y_axis({'min':0,'max':6,'major_gridlines':{'visible':False}})
    ws_ozet.insert_chart('H2',toplu_chart)

    # 2. DERS DETAY SAYFALARI VE 3. ÖZET
    grouped = df.groupby(GROUPING_COLUMNS, dropna=False)
    for group_key, group_df in grouped:
        ders_birim, ders_kodu = str(group_key[0]).strip(), str(group_key[1]).strip()
        if not ders_kodu or ders_kodu.lower()=='nan': continue
        is_uzaktan = "Uzaktan Öğretim" in ders_birim
        suffix = "_UZ" if is_uzaktan else "_OR"
        clean_kodu = ders_kodu.replace('/','_').replace(':','_')
        limit = 31 - len(clean_kodu) - len(suffix) -1
        sheet_name = f"{clean_kodu}{suffix}_{ders_birim[:limit]}".strip('_')
        group_averages = calculate_group_averages(group_df)
        q_averages = calculate_question_averages(group_df)
        n_count = len(group_df)
        gen_avg = calculate_generic_avg(group_df)
        if group_averages:
            summary_list.append({
                "Ders Birim": ders_birim,
                "Ders Kodu": ders_kodu,
                "Kategori": identify_class_and_type(ders_kodu, ders_birim),
                "Katılımcı (N)": n_count,
                **group_averages,
                "Genel Memnuniyet": gen_avg
            })
            avg_df = pd.DataFrame.from_dict(group_averages, orient='index', columns=['Ortalama']).reset_index()
            avg_df.rename(columns={'index':'Grup'}, inplace=True)
            avg_df.loc[len(avg_df)] = ["Memnuniyet Ortalaması", gen_avg]
            avg_df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws_detay = writer.sheets[sheet_name]
            ws_detay.set_column('A:A',25)
            ws_detay.write(len(avg_df)+2,0,f"Katılımcı Sayısı (N): {n_count}")
            chart_perf = create_chart_original_format(workbook, sheet_name, f"$A$2:$A${len(avg_df)+1}", f"$B$2:$B${len(avg_df)+1}", f"{ders_kodu} Memnuniyet Sonuçları")
            ws_detay.insert_chart('D2',chart_perf)
            chart_trend = create_chart_original_format(workbook, sheet_name, f"$J$2:$J${len(q_averages)+1}", f"$K$2:$K${len(q_averages)+1}", "Soru Bazlı Memnuniyet Trendi", is_line=True)
            for i,(q_code,q_val) in enumerate(q_averages.items(),start=2):
                ws_detay.write(i,9,q_code); ws_detay.write(i,10,q_val)
            ws_detay.insert_chart('D18',chart_trend)

    if summary_list:
        summary_df = pd.DataFrame(summary_list)
        summary_cols = ["Ders Birim","Ders Kodu","Kategori","Katılımcı (N)"] + list(QUESTION_GROUPS.keys()) + ["Genel Memnuniyet"]
        summary_df[summary_cols].to_excel(writer, sheet_name='Tum_Dersler_Ozet', index=False)
        ws_all = writer.sheets['Tum_Dersler_Ozet']
        for col_num,_ in enumerate(summary_df[summary_cols].columns.values):
            ws_all.set_column(col_num,col_num,20)

print(f"\nİşlem Tamamlandı. Orijinal grafik formatı korunarak TBMYO raporları oluşturuldu.")