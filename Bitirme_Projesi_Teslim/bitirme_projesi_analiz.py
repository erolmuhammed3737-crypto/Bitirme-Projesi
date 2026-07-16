# -*- coding: utf-8 -*-
"""
TBMYO Bitirme Projesi — İki Dönem Karşılaştırmalı Akademik Analiz
==================================================================

Bu program iki farklı Güz dönemine ait:
  1) Ders memnuniyet anketlerini,
  2) Ders başarı oranı raporlarını
karşılaştırır ve her "Ders Birim" için ayrı Excel dosyaları üretir.

Hocanın istediği temel özellikler:
- Birim aktarma yoktur; kaynak dosyadaki birim kullanılır.
- Birimi boş satırlar "Birimi_Bos.xlsx" dosyasında ayrıca raporlanır.
- 2024-2025 Güz ve 2025-2026 Güz değerleri yan yana gösterilir.
- Her birim dosyasında iki genel özet ve ders detay sayfaları bulunur.
- Memnuniyet sonuçlarında MD_detay.py soru grupları ve grafik biçimi kullanılır.
- Başarı oranında hem kaynak oran hem de devamsız hariç güncel oran bulunur.
- Hocanın gönderdiği "Başarı oranı_Güncel (%)" sütunu varsa formül satır bazında doğrulanır.

Güncel başarı oranı:
    Başarılı Öğrenci Sayısı / (Öğrenci Sayısı - DZ) * 100

Çalıştırma:
    python bitirme_projesi_analiz.py

Dosyalar varsayılan olarak proje içindeki Veriler klasöründen otomatik bulunur.
Başka konumdaki dosyalar için --help komutunu kullanın.
"""

from __future__ import annotations

import argparse
import math
import re
import shutil
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd


# -----------------------------------------------------------------------------
# SABİTLER
# -----------------------------------------------------------------------------
DONEM_1 = "2024-2025 Güz"
DONEM_2 = "2025-2026 Güz"
BIRIMI_BOS = "Birimi Boş"
HEDEF_UST_BIRIM = "Teknik Bilimler Meslek Yüksekokulu"

QUESTION_GROUPS: Dict[str, Tuple[str, ...]] = {
    "Ders İçeriği": ("1_1", "3_1", "14_1"),
    "Öğretim Elemanı": (
        "2_1", "4_1", "5_1", "7_1", "9_1", "10_1", "11_1",
        "12_1", "14_1", "15_1", "16_1",
    ),
    "Ölçme Değerlendirme": ("6_1", "12_1", "13_1", "14_1"),
    "Yöntem": ("4_1", "6_1", "8_1", "10_1", "14_1"),
}
METRIKLER = list(QUESTION_GROUPS) + ["Genel Memnuniyet"]

LIKERT_MAP = {
    "kesinlikle katılmıyorum": 1.0,
    "katılmıyorum": 2.0,
    "pek fazla katılmıyorum": 3.0,
    "katılıyorum": 4.0,
    "biraz katılıyorum": 5.0,
    "tamamen katılıyorum": 6.0,
}
OZEL_SIFIR_6 = "ödev, proje, ekip çalışması, öğrenci sunumları yapılmadı."
OZEL_SIFIR_8 = "ders için kaynak önerilmedi."

BIRIM_ADI_DUZELTMELERI = {
    # Dönemler arasında yalnızca yazım farkı olan program adları.
    "basım ve yayın teknolojileri": "Basım ve Yayım Teknolojileri",
    "basım ve yayınteknolojileri(iö)": "Basım ve Yayım Teknolojileri (İÖ)",
    "basım ve yayın teknolojileri (iö)": "Basım ve Yayım Teknolojileri (İÖ)",
}

RENKLER = {
    "lacivert": "#1F4E79",
    "mavi": "#4F81BD",
    "turuncu": "#ED7D31",
    "acik_mavi": "#D9EAF7",
    "gri": "#D9D9D9",
    "acik_gri": "#F2F2F2",
    "yesil": "#92D050",
    "kirmizi": "#E6B8B7",
    "mor": "#CCC0DA",
    "acik_turuncu": "#FCD5B5",
    "beyaz": "#FFFFFF",
}
GRUP_RENKLERI = {
    "Ders İçeriği": RENKLER["kirmizi"],
    "Öğretim Elemanı": "#B7DEE8",
    "Ölçme Değerlendirme": RENKLER["mor"],
    "Yöntem": RENKLER["acik_turuncu"],
    "Genel Memnuniyet": RENKLER["yesil"],
}


# -----------------------------------------------------------------------------
# GENEL YARDIMCILAR
# -----------------------------------------------------------------------------
def temiz_metin(deger) -> str:
    if deger is None or (isinstance(deger, float) and math.isnan(deger)):
        return ""
    metin = str(deger).replace("\u200b", "").replace("\xa0", " ")
    return re.sub(r"\s+", " ", metin).strip()


def arama_metni(metin: str) -> str:
    metin = temiz_metin(metin).casefold()
    tablo = str.maketrans("çğıöşü", "cgiosu")
    metin = metin.translate(tablo)
    metin = unicodedata.normalize("NFKD", metin)
    return "".join(ch for ch in metin if not unicodedata.combining(ch))


def grup_no_temizle(deger) -> str:
    metin = temiz_metin(deger)
    return re.sub(r"\.0$", "", metin)


def birim_temizle(deger) -> str:
    metin = temiz_metin(deger)
    if not metin or metin.casefold() in {"nan", "none"}:
        return BIRIMI_BOS
    anahtar = arama_metni(metin)
    anahtar_bos = anahtar.replace(" ", "")
    if anahtar_bos in {"basimveyayinteknolojileri", "basimveyayimteknolojileri"}:
        return "Basım ve Yayım Teknolojileri"
    if anahtar_bos in {
        "basimveyayinteknolojileri(io)", "basimveyayimteknolojileri(io)",
        "basimveyayinteknolojileri(i.o.)", "basimveyayimteknolojileri(i.o.)",
    }:
        return "Basım ve Yayım Teknolojileri (İÖ)"
    return BIRIM_ADI_DUZELTMELERI.get(metin.casefold(), metin)


def guvenli_dosya_adi(metin: str, limit: int = 120) -> str:
    metin = temiz_metin(metin).replace("İ", "I").replace("ı", "i")
    metin = re.sub(r"[<>:\"/\\|?*]", "-", metin)
    metin = re.sub(r"\s+", "_", metin).strip(" ._")
    return (metin or "Adsiz")[:limit]


def guvenli_sayfa_adi(metin: str, kullanilan: set[str]) -> str:
    aday = re.sub(r"[\[\]:*?/\\]", "-", temiz_metin(metin))[:31] or "Sayfa"
    temel = aday
    sayac = 2
    while aday.casefold() in {x.casefold() for x in kullanilan}:
        ek = f"_{sayac}"
        aday = f"{temel[:31-len(ek)]}{ek}"
        sayac += 1
    kullanilan.add(aday)
    return aday


def benzersiz_birlestir(degerler: Iterable) -> str:
    sonuc: List[str] = []
    for deger in degerler:
        metin = temiz_metin(deger)
        if metin and metin.casefold() not in {x.casefold() for x in sonuc}:
            sonuc.append(metin)
    return " / ".join(sonuc)


def sayisal(seri: pd.Series) -> pd.Series:
    return pd.to_numeric(seri, errors="coerce")


def yuvarla(deger, basamak: int = 2):
    if deger is None or pd.isna(deger):
        return np.nan
    return round(float(deger), basamak)


def excel_deger(deger):
    """XlsxWriter doğrudan NaN yazamadığı için eksik sayıları boş hücreye çevirir."""
    if deger is None or pd.isna(deger):
        return ""
    return deger


def fark(yeni, eski):
    if pd.isna(yeni) or pd.isna(eski):
        return np.nan
    return round(float(yeni) - float(eski), 2)


def ders_sinifi(ders_kodu: str) -> str:
    kod = temiz_metin(ders_kodu)
    eslesme = re.search(r"[A-Za-zÇĞİÖŞÜçğıöşü]+\s*(\d)", kod)
    if not eslesme:
        return "Sınıf Bilinmiyor"
    rakam = eslesme.group(1)
    if rakam in {"1", "2", "3", "4"}:
        return f"{rakam}. Sınıf"
    return "Sınıf Bilinmiyor"


def ogretim_turu(satir: pd.Series) -> str:
    alanlar = [
        temiz_metin(satir.get("Ders Birim", "")),
        temiz_metin(satir.get("Ders Alt Birim", "")),
        temiz_metin(satir.get("Ders Üst Birim", "")),
    ]
    metin = " ".join(alanlar).casefold()
    if "uzaktan" in metin:
        return "Uzaktan Öğretim"
    if "(iö)" in metin or "ikinci öğretim" in metin:
        return "İkinci Öğretim"
    return "Birinci Öğretim"


def kategori_uret(satir: pd.Series) -> str:
    return f"{ders_sinifi(satir.get('Ders Kodu', ''))} - {ogretim_turu(satir)}"


def soru_kodu(sutun: str) -> Optional[str]:
    eslesme = re.match(r"^\s*(\d+_1)\b", str(sutun))
    return eslesme.group(1) if eslesme else None


def soru_sirasi(kod: str) -> int:
    try:
        return int(str(kod).split("_")[0])
    except (ValueError, IndexError):
        return 999


def likert_donustur(deger, kod: str):
    if pd.isna(deger):
        return np.nan
    if isinstance(deger, (int, float, np.number)):
        sayi = float(deger)
        return sayi if 0 <= sayi <= 6 else np.nan
    metin = temiz_metin(deger).casefold()
    if kod == "6_1" and metin == OZEL_SIFIR_6:
        return 0.0
    if kod == "8_1" and metin == OZEL_SIFIR_8:
        return 0.0
    return LIKERT_MAP.get(metin, np.nan)


def agirlikli_ortalama(seri: pd.Series, agirlik: pd.Series) -> float:
    s = sayisal(seri)
    a = sayisal(agirlik)
    gecerli = s.notna() & a.notna() & (a > 0)
    if not gecerli.any():
        return np.nan
    return float(np.average(s[gecerli], weights=a[gecerli]))


# -----------------------------------------------------------------------------
# DOSYA BULMA VE DOĞRULAMA
# -----------------------------------------------------------------------------
@dataclass
class GirdiDosyalari:
    memnuniyet_1: Path
    memnuniyet_2: Path
    basari_1: Path
    basari_2: Path


def otomatik_dosya_bul(kok: Path) -> GirdiDosyalari:
    tercih = kok / "Veriler"
    arama_koku = tercih if tercih.exists() else kok
    dosyalar = [p for p in arama_koku.rglob("*.xlsx") if "Sonuclar" not in p.parts]

    def bul(*parcalar: str) -> Path:
        bulunan = []
        for yol in dosyalar:
            ad = arama_metni(yol.name)
            if all(arama_metni(p) in ad for p in parcalar):
                bulunan.append(yol)
        if len(bulunan) == 1:
            return bulunan[0]
        if not bulunan:
            raise FileNotFoundError(
                "Şu anahtarlarla eşleşen Excel bulunamadı: " + ", ".join(parcalar)
            )
        # En kısa/özgül adı seç; çıktı dosyaları dışarıda bırakılmıştır.
        bulunan.sort(key=lambda p: (len(p.name), str(p)))
        return bulunan[0]

    # Teslim paketindeki sade adlar önce denenir.
    sabitler = {
        "memnuniyet_1": arama_koku / "memnuniyet_2024_2025_guz.xlsx",
        "memnuniyet_2": arama_koku / "memnuniyet_2025_2026_guz.xlsx",
        "basari_1": arama_koku / "basari_2024_2025_guz.xlsx",
        "basari_2": arama_koku / "basari_2025_2026_guz.xlsx",
    }
    return GirdiDosyalari(
        memnuniyet_1=sabitler["memnuniyet_1"] if sabitler["memnuniyet_1"].exists() else bul("2024", "2025", "guz", "anket"),
        memnuniyet_2=sabitler["memnuniyet_2"] if sabitler["memnuniyet_2"].exists() else bul("2025", "2026", "guz", "anket"),
        basari_1=sabitler["basari_1"] if sabitler["basari_1"].exists() else bul("2024", "2025", "guz", "basari"),
        basari_2=sabitler["basari_2"] if sabitler["basari_2"].exists() else bul("2025", "2026", "guz", "basari"),
    )


def zorunlu_sutun_kontrol(df: pd.DataFrame, sutunlar: Sequence[str], dosya: Path):
    eksik = [s for s in sutunlar if s not in df.columns]
    if eksik:
        raise ValueError(
            f"{dosya.name} dosyasında zorunlu sütunlar eksik: {', '.join(eksik)}"
        )


def excel_oku(yol: Path) -> pd.DataFrame:
    if not yol.exists():
        raise FileNotFoundError(f"Dosya bulunamadı: {yol}")
    df = pd.read_excel(yol, sheet_name=0)
    df.columns = [temiz_metin(c) for c in df.columns]
    return df


# -----------------------------------------------------------------------------
# MEMNUNİYET VERİSİ
# -----------------------------------------------------------------------------
def memnuniyet_hazirla(yol: Path) -> Tuple[pd.DataFrame, Dict[str, str]]:
    df = excel_oku(yol)
    zorunlu_sutun_kontrol(
        df,
        ["Ders Üst Birim", "Ders Birim", "Ders Kodu", "Grup No", "Öğretim Üyesi"],
        yol,
    )

    # Sütun adındaki uzun soru metnini koruyup hesaplamada kısa kod kullanıyoruz.
    soru_metinleri: Dict[str, str] = {}
    yeniden_adlandir = {}
    for sutun in df.columns:
        kod = soru_kodu(sutun)
        if kod:
            soru_metinleri[kod] = temiz_metin(sutun[len(kod):])
            yeniden_adlandir[sutun] = kod
    df = df.rename(columns=yeniden_adlandir)
    soru_kodlari = sorted(set(yeniden_adlandir.values()), key=soru_sirasi)
    if not soru_kodlari:
        raise ValueError(f"{yol.name} içinde 1_1, 2_1 ... biçiminde soru sütunu yok.")

    for sutun in ["Ders Üst Birim", "Ders Birim", "Ders Alt Birim", "Ders Kodu", "Öğretim Üyesi", "Ders Adı"]:
        if sutun not in df.columns:
            df[sutun] = ""
        df[sutun] = df[sutun].map(temiz_metin)
    df["Ders Birim"] = df["Ders Birim"].map(birim_temizle)
    df["Grup No"] = df["Grup No"].map(grup_no_temizle)
    df = df[df["Ders Kodu"].map(temiz_metin) != ""].copy()

    for kod in soru_kodlari:
        df[kod] = df[kod].map(lambda v, q=kod: likert_donustur(v, q))

    df["Kategori"] = df.apply(kategori_uret, axis=1)
    df["Ders Anahtarı"] = df["Ders Kodu"] + "|" + df["Grup No"]
    df.attrs["soru_kodlari"] = soru_kodlari
    return df, soru_metinleri


def memnuniyete_ders_adi_ekle(memnuniyet: pd.DataFrame, basari: pd.DataFrame) -> pd.DataFrame:
    """Başarı dosyasındaki Ders Adı bilgisini kod+grup üzerinden ankete ekler.

    Memnuniyet kaynaklarında Ders Adı sütunu bulunmadığı için hocanın istediği
    grafik başlığı ve özet tablosu, aynı dönemin başarı dosyasından zenginleştirilir.
    Eşleşme yoksa ders kodu kullanılmaya devam eder.
    """
    soru_kodlari = list(memnuniyet.attrs.get("soru_kodlari", []))
    harita = (
        basari.groupby(["Ders Kodu", "Grup No"], dropna=False)["Ders Adı"]
        .apply(benzersiz_birlestir)
        .to_dict()
    )
    sonuc = memnuniyet.copy()
    sonuc["Ders Adı"] = [
        harita.get((kod, grup), temiz_metin(mevcut))
        for kod, grup, mevcut in zip(sonuc["Ders Kodu"], sonuc["Grup No"], sonuc["Ders Adı"])
    ]
    sonuc.attrs["soru_kodlari"] = soru_kodlari
    return sonuc


def metrik_hesapla(df: pd.DataFrame, soru_kodlari: Sequence[str]) -> Dict[str, float]:
    sonuc: Dict[str, float] = {}
    for grup, kodlar in QUESTION_GROUPS.items():
        mevcut = [q for q in kodlar if q in soru_kodlari and q in df.columns]
        if not mevcut:
            sonuc[grup] = np.nan
            continue
        dizi = df[mevcut].to_numpy(dtype=float)
        gecerli = np.isfinite(dizi)
        sonuc[grup] = float(np.nansum(dizi) / gecerli.sum()) if gecerli.sum() else np.nan
    mevcut_tum = [q for q in soru_kodlari if q in df.columns]
    dizi = df[mevcut_tum].to_numpy(dtype=float)
    gecerli = np.isfinite(dizi)
    sonuc["Genel Memnuniyet"] = float(np.nansum(dizi) / gecerli.sum()) if gecerli.sum() else np.nan
    return sonuc


def soru_ortalamalari(df: pd.DataFrame, soru_kodlari: Sequence[str]) -> Dict[str, float]:
    return {q: float(df[q].mean()) if df[q].notna().any() else np.nan for q in soru_kodlari}


def memnuniyet_ders_ozeti(df: pd.DataFrame) -> pd.DataFrame:
    soru_kodlari: List[str] = list(df.attrs.get("soru_kodlari", []))
    satirlar = []
    for (kod, grup), alt in df.groupby(["Ders Kodu", "Grup No"], dropna=False, sort=True):
        metrik = metrik_hesapla(alt, soru_kodlari)
        satirlar.append({
            "Ders Kodu": temiz_metin(kod),
            "Grup No": grup_no_temizle(grup),
            "Kategori": benzersiz_birlestir(alt["Kategori"]),
            "Ders Adı": benzersiz_birlestir(alt["Ders Adı"]),
            "Öğretim Üyesi": benzersiz_birlestir(alt["Öğretim Üyesi"]),
            "Katılımcı (N)": int(len(alt)),
            **{m: yuvarla(metrik[m]) for m in METRIKLER},
        })
    kolonlar = ["Ders Kodu", "Grup No", "Kategori", "Ders Adı", "Öğretim Üyesi", "Katılımcı (N)"] + METRIKLER
    return pd.DataFrame(satirlar, columns=kolonlar)


def memnuniyet_karsilastirma(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
    o1 = memnuniyet_ders_ozeti(df1).rename(columns={
        "Kategori": f"Kategori {DONEM_1}",
        "Ders Adı": f"Ders Adı {DONEM_1}",
        "Öğretim Üyesi": f"Öğretim Üyesi {DONEM_1}",
        "Katılımcı (N)": f"Katılımcı {DONEM_1}",
        **{m: f"{m} {DONEM_1}" for m in METRIKLER},
    })
    o2 = memnuniyet_ders_ozeti(df2).rename(columns={
        "Kategori": f"Kategori {DONEM_2}",
        "Ders Adı": f"Ders Adı {DONEM_2}",
        "Öğretim Üyesi": f"Öğretim Üyesi {DONEM_2}",
        "Katılımcı (N)": f"Katılımcı {DONEM_2}",
        **{m: f"{m} {DONEM_2}" for m in METRIKLER},
    })
    tablo = o1.merge(o2, on=["Ders Kodu", "Grup No"], how="outer")
    tablo["Kategori"] = tablo[f"Kategori {DONEM_2}"].fillna(tablo[f"Kategori {DONEM_1}"])
    for m in METRIKLER:
        tablo[f"{m} Fark"] = tablo.apply(
            lambda r, met=m: fark(r.get(f"{met} {DONEM_2}"), r.get(f"{met} {DONEM_1}")),
            axis=1,
        )
    kolonlar = [
        "Ders Kodu", "Grup No", "Kategori",
        f"Ders Adı {DONEM_1}", f"Ders Adı {DONEM_2}",
        f"Öğretim Üyesi {DONEM_1}", f"Öğretim Üyesi {DONEM_2}",
        f"Katılımcı {DONEM_1}", f"Katılımcı {DONEM_2}",
    ]
    for m in METRIKLER:
        kolonlar.extend([f"{m} {DONEM_1}", f"{m} {DONEM_2}", f"{m} Fark"])
    tablo = tablo[kolonlar]
    return tablo.sort_values(["Kategori", "Ders Kodu", "Grup No"], na_position="last").reset_index(drop=True)


def memnuniyet_birim_genel(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
    soru1 = list(df1.attrs.get("soru_kodlari", []))
    soru2 = list(df2.attrs.get("soru_kodlari", []))
    kategoriler = sorted(set(df1["Kategori"].dropna()) | set(df2["Kategori"].dropna()))
    satirlar = []
    for kategori in kategoriler + ["BÖLÜM GENELİ"]:
        a1 = df1 if kategori == "BÖLÜM GENELİ" else df1[df1["Kategori"] == kategori]
        a2 = df2 if kategori == "BÖLÜM GENELİ" else df2[df2["Kategori"] == kategori]
        m1 = metrik_hesapla(a1, soru1) if not a1.empty else {m: np.nan for m in METRIKLER}
        m2 = metrik_hesapla(a2, soru2) if not a2.empty else {m: np.nan for m in METRIKLER}
        satir = {
            "Kategori": kategori,
            f"Katılımcı {DONEM_1}": int(len(a1)),
            f"Katılımcı {DONEM_2}": int(len(a2)),
        }
        for m in METRIKLER:
            satir[f"{m} {DONEM_1}"] = yuvarla(m1[m])
            satir[f"{m} {DONEM_2}"] = yuvarla(m2[m])
            satir[f"{m} Fark"] = fark(m2[m], m1[m])
        satirlar.append(satir)
    return pd.DataFrame(satirlar)


# -----------------------------------------------------------------------------
# BAŞARI VERİSİ
# -----------------------------------------------------------------------------
def basari_hazirla(yol: Path, ust_birim: str) -> pd.DataFrame:
    df = excel_oku(yol)
    zorunlu_sutun_kontrol(
        df,
        [
            "Üst Birim", "Birim", "Ders Kodu", "Grup No", "Ders Adı", "Öğretim Üyesi",
            "Öğrenci Sayısı", "Başarılı Öğrenci Sayısı", "Başarısız Öğrenci Sayısı", "DZ",
        ],
        yol,
    )
    # Kaynak dosyada hem eski başarı oranı hem de hocanın eklediği
    # "Başarı oranı_Güncel (%)" alanı bulunabilir. Eski oranı raporda
    # koruyor, güncel alanı ise yeniden hesaplanan formülü doğrulamak için
    # ayrıca okuyoruz.
    def oran_anahtari(sutun: str) -> str:
        return re.sub(r"[^a-z0-9]", "", arama_metni(sutun))

    guncel_oran_sutunu = next(
        (c for c in df.columns
         if "basariorani" in oran_anahtari(c) and "guncel" in oran_anahtari(c)),
        None,
    )
    oran_sutunu = next(
        (c for c in df.columns
         if "basariorani" in oran_anahtari(c) and "guncel" not in oran_anahtari(c)),
        None,
    )
    if oran_sutunu is None:
        raise ValueError(f"{yol.name} içinde başarı oranı sütunu bulunamadı.")

    yeniden_adlandir = {oran_sutunu: "Kaynak Başarı Oranı"}
    if guncel_oran_sutunu is not None:
        yeniden_adlandir[guncel_oran_sutunu] = "Kaynak Başarı Oranı Güncel"
    df = df.rename(columns=yeniden_adlandir)
    if "Kaynak Başarı Oranı Güncel" not in df.columns:
        df["Kaynak Başarı Oranı Güncel"] = np.nan

    for sutun in ["Üst Birim", "Birim", "Ders Kodu", "Ders Adı", "Öğretim Üyesi"]:
        df[sutun] = df[sutun].map(temiz_metin)
    df["Grup No"] = df["Grup No"].map(grup_no_temizle)
    df = df[df["Ders Kodu"] != ""].copy()
    hedef = arama_metni(ust_birim)
    df = df[df["Üst Birim"].map(arama_metni) == hedef].copy()
    df["Birim"] = df["Birim"].map(birim_temizle)

    for sutun in [
        "Öğrenci Sayısı", "Başarılı Öğrenci Sayısı", "Başarısız Öğrenci Sayısı",
        "Kaynak Başarı Oranı", "Kaynak Başarı Oranı Güncel", "DZ",
    ]:
        df[sutun] = sayisal(df[sutun])
    for sutun in [
        "Öğrenci Sayısı", "Başarılı Öğrenci Sayısı", "Başarısız Öğrenci Sayısı",
        "Kaynak Başarı Oranı", "DZ",
    ]:
        df[sutun] = df[sutun].fillna(0)

    # Hocanın gönderdiği güncel sütunu varsa formülü satır bazında doğrula.
    # Payda sıfır olduğunda Excel'deki #DIV/0! değeri sayıya dönüşmez ve
    # bilinçli olarak karşılaştırma dışında tutulur.
    payda = df["Öğrenci Sayısı"] - df["DZ"]
    hesaplanan = pd.Series(np.nan, index=df.index, dtype=float)
    gecerli_payda = payda > 0
    hesaplanan.loc[gecerli_payda] = (
        df.loc[gecerli_payda, "Başarılı Öğrenci Sayısı"]
        / payda.loc[gecerli_payda] * 100.0
    )
    resmi = df["Kaynak Başarı Oranı Güncel"]
    karsilastirilabilir = gecerli_payda & resmi.notna()
    uyusmaz = karsilastirilabilir & ((resmi - hesaplanan).abs() > 0.01)
    df.attrs["guncel_oran_dogrulama"] = {
        "resmi_sutun_var": guncel_oran_sutunu is not None,
        "karsilastirilan_satir": int(karsilastirilabilir.sum()),
        "sifir_payda_satir": int((~gecerli_payda).sum()),
        "uyusmaz_satir": int(uyusmaz.sum()),
    }
    if uyusmaz.any():
        print(
            f"  UYARI: {yol.name} içinde hocanın güncel oran sütunu ile "
            f"yeniden hesaplanan oran arasında {int(uyusmaz.sum())} satır fark var."
        )

    # Başarı dosyasında kategori için Ders Birim adı kullanılır.
    df["Ders Birim"] = df["Birim"]
    df["Ders Alt Birim"] = ""
    df["Ders Üst Birim"] = df["Üst Birim"]
    df["Kategori"] = df.apply(kategori_uret, axis=1)
    return df


def basari_ders_ozeti(df: pd.DataFrame) -> pd.DataFrame:
    satirlar = []
    for (kod, grup), alt in df.groupby(["Ders Kodu", "Grup No"], dropna=False, sort=True):
        ogr = float(alt["Öğrenci Sayısı"].sum())
        bas = float(alt["Başarılı Öğrenci Sayısı"].sum())
        basarisiz = float(alt["Başarısız Öğrenci Sayısı"].sum())
        dz = float(alt["DZ"].sum())
        # Kaynak oran birden fazla satır varsa öğrenci sayısına göre ağırlıklandırılır.
        kaynak_oran = agirlikli_ortalama(alt["Kaynak Başarı Oranı"], alt["Öğrenci Sayısı"])
        guncel_payda = ogr - dz
        guncel_oran = (bas / guncel_payda * 100.0) if guncel_payda > 0 else np.nan
        satirlar.append({
            "Ders Kodu": temiz_metin(kod),
            "Grup No": grup_no_temizle(grup),
            "Kategori": benzersiz_birlestir(alt["Kategori"]),
            "Ders Adı": benzersiz_birlestir(alt["Ders Adı"]),
            "Öğretim Üyesi": benzersiz_birlestir(alt["Öğretim Üyesi"]),
            "Öğrenci Sayısı": int(ogr),
            "Başarılı Öğrenci Sayısı": int(bas),
            "Başarısız Öğrenci Sayısı": int(basarisiz),
            "DZ": int(dz),
            "Devamsız Hariç Öğrenci Sayısı": int(max(guncel_payda, 0)),
            "Başarı Oranı": yuvarla(kaynak_oran),
            "Başarı Oranı Güncel": yuvarla(guncel_oran),
        })
    kolonlar = [
        "Ders Kodu", "Grup No", "Kategori", "Ders Adı", "Öğretim Üyesi",
        "Öğrenci Sayısı", "Başarılı Öğrenci Sayısı", "Başarısız Öğrenci Sayısı",
        "DZ", "Devamsız Hariç Öğrenci Sayısı", "Başarı Oranı", "Başarı Oranı Güncel",
    ]
    return pd.DataFrame(satirlar, columns=kolonlar)


def basari_karsilastirma(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
    o1 = basari_ders_ozeti(df1).rename(columns={
        c: f"{c} {DONEM_1}" for c in [
            "Ders Adı", "Öğretim Üyesi", "Öğrenci Sayısı", "Başarılı Öğrenci Sayısı",
            "Başarısız Öğrenci Sayısı", "DZ", "Devamsız Hariç Öğrenci Sayısı",
            "Başarı Oranı", "Başarı Oranı Güncel",
        ]
    }).rename(columns={"Kategori": f"Kategori {DONEM_1}"})
    o2 = basari_ders_ozeti(df2).rename(columns={
        c: f"{c} {DONEM_2}" for c in [
            "Ders Adı", "Öğretim Üyesi", "Öğrenci Sayısı", "Başarılı Öğrenci Sayısı",
            "Başarısız Öğrenci Sayısı", "DZ", "Devamsız Hariç Öğrenci Sayısı",
            "Başarı Oranı", "Başarı Oranı Güncel",
        ]
    }).rename(columns={"Kategori": f"Kategori {DONEM_2}"})
    tablo = o1.merge(o2, on=["Ders Kodu", "Grup No"], how="outer")
    tablo["Kategori"] = tablo[f"Kategori {DONEM_2}"].fillna(tablo[f"Kategori {DONEM_1}"])
    tablo["Güncel Başarı Oranı Fark"] = tablo.apply(
        lambda r: fark(r.get(f"Başarı Oranı Güncel {DONEM_2}"), r.get(f"Başarı Oranı Güncel {DONEM_1}")),
        axis=1,
    )
    kolonlar = ["Ders Kodu", "Grup No", "Kategori"]
    for bilgi in ["Ders Adı", "Öğretim Üyesi"]:
        kolonlar += [f"{bilgi} {DONEM_1}", f"{bilgi} {DONEM_2}"]
    for metrik in [
        "Öğrenci Sayısı", "Başarılı Öğrenci Sayısı", "Başarısız Öğrenci Sayısı", "DZ",
        "Devamsız Hariç Öğrenci Sayısı", "Başarı Oranı", "Başarı Oranı Güncel",
    ]:
        kolonlar += [f"{metrik} {DONEM_1}", f"{metrik} {DONEM_2}"]
    kolonlar.append("Güncel Başarı Oranı Fark")
    return tablo[kolonlar].sort_values(["Kategori", "Ders Kodu", "Grup No"]).reset_index(drop=True)


def basari_birim_genel(df1: pd.DataFrame, df2: pd.DataFrame) -> pd.DataFrame:
    kategoriler = sorted(set(df1["Kategori"].dropna()) | set(df2["Kategori"].dropna()))
    satirlar = []
    for kategori in kategoriler + ["BÖLÜM GENELİ"]:
        satir = {"Kategori": kategori}
        for donem, veri in [(DONEM_1, df1), (DONEM_2, df2)]:
            alt = veri if kategori == "BÖLÜM GENELİ" else veri[veri["Kategori"] == kategori]
            ogr = float(alt["Öğrenci Sayısı"].sum()) if not alt.empty else 0
            bas = float(alt["Başarılı Öğrenci Sayısı"].sum()) if not alt.empty else 0
            basarisiz = float(alt["Başarısız Öğrenci Sayısı"].sum()) if not alt.empty else 0
            dz = float(alt["DZ"].sum()) if not alt.empty else 0
            kaynak = agirlikli_ortalama(alt["Kaynak Başarı Oranı"], alt["Öğrenci Sayısı"]) if not alt.empty else np.nan
            payda = ogr - dz
            guncel = bas / payda * 100 if payda > 0 else np.nan
            satir[f"Öğrenci Sayısı {donem}"] = int(ogr)
            satir[f"Başarılı {donem}"] = int(bas)
            satir[f"Başarısız {donem}"] = int(basarisiz)
            satir[f"DZ {donem}"] = int(dz)
            satir[f"Başarı Oranı {donem}"] = yuvarla(kaynak)
            satir[f"Başarı Oranı Güncel {donem}"] = yuvarla(guncel)
        satir["Güncel Başarı Oranı Fark"] = fark(
            satir[f"Başarı Oranı Güncel {DONEM_2}"], satir[f"Başarı Oranı Güncel {DONEM_1}"]
        )
        satirlar.append(satir)
    return pd.DataFrame(satirlar)


# -----------------------------------------------------------------------------
# EXCEL BİÇİMLENDİRME
# -----------------------------------------------------------------------------
def excel_formatlari(workbook):
    return {
        "baslik": workbook.add_format({
            "bold": True, "font_color": "white", "bg_color": RENKLER["lacivert"],
            "align": "center", "valign": "vcenter", "font_size": 13, "border": 1,
        }),
        "alt_baslik": workbook.add_format({
            "bold": True, "font_color": "white", "bg_color": RENKLER["mavi"],
            "align": "center", "valign": "vcenter", "text_wrap": True, "border": 1,
        }),
        "normal": workbook.add_format({"border": 1, "valign": "vcenter"}),
        "normal_ortala": workbook.add_format({"border": 1, "align": "center", "valign": "vcenter"}),
        "sayi": workbook.add_format({"border": 1, "align": "center", "num_format": "0"}),
        "ondalik": workbook.add_format({"border": 1, "align": "center", "num_format": "0.00"}),
        "yuzde": workbook.add_format({"border": 1, "align": "center", "num_format": "0.00"}),
        "fark": workbook.add_format({"border": 1, "align": "center", "num_format": "+0.00;-0.00;0.00"}),
        "aciklama": workbook.add_format({"text_wrap": True, "valign": "top", "bg_color": "#FFF2CC", "border": 1}),
        "toplam": workbook.add_format({"bold": True, "bg_color": "#FCE4D6", "border": 1, "align": "center"}),
    }


def tablo_bicimlendir(writer, sayfa: str, df: pd.DataFrame, baslik: str, fark_kolonlari: Iterable[str] = ()): 
    ws = writer.sheets[sayfa]
    wb = writer.book
    fmt = excel_formatlari(wb)
    son_kolon = max(len(df.columns) - 1, 0)
    ws.set_row(0, 28)
    ws.merge_range(0, 0, 0, son_kolon, baslik, fmt["baslik"])
    for i, kolon in enumerate(df.columns):
        ws.write(1, i, kolon, fmt["alt_baslik"])
        metin = str(kolon)
        if any(x in metin for x in ["Öğretim Üyesi", "Ders Adı"]):
            ws.set_column(i, i, 28)
        elif metin in {"Ders Kodu", "Grup No"}:
            ws.set_column(i, i, 12)
        elif "Kategori" in metin:
            ws.set_column(i, i, 28)
        elif "Katılımcı" in metin or "Sayısı" in metin or metin.startswith("DZ"):
            ws.set_column(i, i, 14)
        else:
            ws.set_column(i, i, 18)
    ws.freeze_panes(2, 3)
    ws.autofilter(1, 0, 1 + len(df), son_kolon)
    ws.hide_gridlines(2)
    ws.set_default_row(20)
    fark_set = set(fark_kolonlari)
    for col_idx, kolon in enumerate(df.columns):
        if kolon in fark_set:
            ws.conditional_format(2, col_idx, 1 + len(df), col_idx, {
                "type": "3_color_scale", "min_color": "#F8696B", "mid_color": "#FFEB84", "max_color": "#63BE7B"
            })


def dataframe_yaz(writer, sayfa: str, df: pd.DataFrame, baslik: str, fark_kolonlari: Iterable[str] = ()): 
    # Başlık için ilk satırı boş bırakıp tabloyu ikinci satırdan başlatıyoruz.
    df.to_excel(writer, sheet_name=sayfa, index=False, startrow=1)
    tablo_bicimlendir(writer, sayfa, df, baslik, fark_kolonlari)


def karsilastirma_grafigi(workbook, sheet_name: str, ilk_veri_satiri: int, son_veri_satiri: int,
                         kategori_col: int, eski_col: int, yeni_col: int, baslik: str):
    chart = workbook.add_chart({"type": "column"})
    chart.add_series({
        "name": DONEM_1,
        "categories": [sheet_name, ilk_veri_satiri, kategori_col, son_veri_satiri, kategori_col],
        "values": [sheet_name, ilk_veri_satiri, eski_col, son_veri_satiri, eski_col],
        "fill": {"color": RENKLER["mavi"]},
        "data_labels": {"value": True, "num_format": "0.00"},
    })
    chart.add_series({
        "name": DONEM_2,
        "categories": [sheet_name, ilk_veri_satiri, kategori_col, son_veri_satiri, kategori_col],
        "values": [sheet_name, ilk_veri_satiri, yeni_col, son_veri_satiri, yeni_col],
        "fill": {"color": RENKLER["turuncu"]},
        "data_labels": {"value": True, "num_format": "0.00"},
    })
    chart.set_title({"name": baslik})
    chart.set_y_axis({"min": 0, "max": 6, "major_gridlines": {"visible": False}})
    chart.set_x_axis({"major_gridlines": {"visible": False}})
    chart.set_legend({"position": "bottom"})
    chart.set_style(10)
    return chart


def memnuniyet_detay_sayfasi(writer, sayfa: str, birim: str, kod: str, grup: str,
                              d1: pd.DataFrame, d2: pd.DataFrame,
                              soru_kodlari: Sequence[str], soru_metinleri: Dict[str, str]):
    wb = writer.book
    ws = wb.add_worksheet(sayfa)
    writer.sheets[sayfa] = ws
    fmt = excel_formatlari(wb)
    ws.hide_gridlines(2)
    ws.set_column("A:A", 25)
    ws.set_column("B:D", 18)
    ws.set_column("F:F", 10)
    ws.set_column("G:G", 58)
    ws.set_column("H:J", 16)

    ders_adi_degerleri = []
    for veri in (d1, d2):
        if "Ders Adı" in veri.columns:
            ders_adi_degerleri.extend(veri["Ders Adı"].tolist())
    ders_adi = benzersiz_birlestir(ders_adi_degerleri)
    ana_baslik = f"{kod} / Grup {grup}"
    if ders_adi:
        ana_baslik += f" — {ders_adi}"
    ana_baslik += f" — {birim}"
    ws.merge_range("A1:J1", ana_baslik, fmt["baslik"])
    ws.write("A2", "Öğretim Üyesi", fmt["alt_baslik"])
    ws.write("B2", benzersiz_birlestir(d1["Öğretim Üyesi"]), fmt["normal"])
    ws.write("C2", DONEM_1, fmt["alt_baslik"])
    ws.write("D2", len(d1), fmt["sayi"])
    ws.write("A3", "Öğretim Üyesi", fmt["alt_baslik"])
    ws.write("B3", benzersiz_birlestir(d2["Öğretim Üyesi"]), fmt["normal"])
    ws.write("C3", DONEM_2, fmt["alt_baslik"])
    ws.write("D3", len(d2), fmt["sayi"])

    m1 = metrik_hesapla(d1, soru_kodlari) if not d1.empty else {m: np.nan for m in METRIKLER}
    m2 = metrik_hesapla(d2, soru_kodlari) if not d2.empty else {m: np.nan for m in METRIKLER}
    ws.write_row(4, 0, ["Kategori", DONEM_1, DONEM_2, "Fark"], fmt["alt_baslik"])
    for i, m in enumerate(METRIKLER, start=5):
        grup_fmt = wb.add_format({"border": 1, "bold": True, "bg_color": GRUP_RENKLERI[m]})
        ws.write(i, 0, m, grup_fmt)
        ws.write(i, 1, excel_deger(yuvarla(m1[m])), fmt["ondalik"])
        ws.write(i, 2, excel_deger(yuvarla(m2[m])), fmt["ondalik"])
        ws.write(i, 3, excel_deger(fark(m2[m], m1[m])), fmt["fark"])

    grafik_basligi = f"{kod}" + (f" - {ders_adi}" if ders_adi else "") + " Öğrenci Memnuniyet Sonuçları"
    chart = karsilastirma_grafigi(wb, sayfa, 5, 5 + len(METRIKLER) - 1, 0, 1, 2,
                                  grafik_basligi)
    ws.insert_chart("A13", chart, {"x_scale": 1.25, "y_scale": 1.15})

    q1 = soru_ortalamalari(d1, soru_kodlari) if not d1.empty else {q: np.nan for q in soru_kodlari}
    q2 = soru_ortalamalari(d2, soru_kodlari) if not d2.empty else {q: np.nan for q in soru_kodlari}
    ws.write_row(4, 5, ["Soru", "Soru Metni", DONEM_1, DONEM_2, "Fark"], fmt["alt_baslik"])
    for i, q in enumerate(soru_kodlari, start=5):
        ws.write(i, 5, q, fmt["normal_ortala"])
        ws.write(i, 6, soru_metinleri.get(q, ""), fmt["normal"])
        ws.write(i, 7, excel_deger(yuvarla(q1[q])), fmt["ondalik"])
        ws.write(i, 8, excel_deger(yuvarla(q2[q])), fmt["ondalik"])
        ws.write(i, 9, excel_deger(fark(q2[q], q1[q])), fmt["fark"])

    trend = wb.add_chart({"type": "line"})
    son = 5 + len(soru_kodlari) - 1
    for ad, col, renk in [(DONEM_1, 7, RENKLER["mavi"]), (DONEM_2, 8, RENKLER["turuncu"])]:
        trend.add_series({
            "name": ad,
            "categories": [sayfa, 5, 5, son, 5],
            "values": [sayfa, 5, col, son, col],
            "line": {"color": renk, "width": 2.25},
            "marker": {"type": "circle", "size": 5},
        })
    trend.set_title({"name": "Soru Bazlı Memnuniyet Trendi"})
    trend.set_y_axis({"min": 0, "max": 6, "major_gridlines": {"visible": False}})
    trend.set_legend({"position": "bottom"})
    ws.insert_chart("F23", trend, {"x_scale": 1.25, "y_scale": 1.15})
    ws.freeze_panes(5, 0)


def memnuniyet_workbook_yaz(yol: Path, birim: str, d1: pd.DataFrame, d2: pd.DataFrame,
                             soru_metinleri: Dict[str, str]):
    soru_kodlari = sorted(
        set(d1.attrs.get("soru_kodlari", [])) | set(d2.attrs.get("soru_kodlari", [])),
        key=soru_sirasi,
    )
    with pd.ExcelWriter(yol, engine="xlsxwriter") as writer:
        ozet = memnuniyet_karsilastirma(d1, d2)
        farklar = [c for c in ozet.columns if c.endswith(" Fark")]
        dataframe_yaz(writer, "01_Tum_Dersler_Ozet", ozet,
                       f"{birim} — Tüm Dersler Karşılaştırmalı Memnuniyet Özeti", farklar)

        genel = memnuniyet_birim_genel(d1, d2)
        genel_fark = [c for c in genel.columns if c.endswith(" Fark")]
        dataframe_yaz(writer, "02_Bolum_Genel_Analiz", genel,
                       f"{birim} — Bölüm Genel Analizi", genel_fark)
        ws_genel = writer.sheets["02_Bolum_Genel_Analiz"]
        wb = writer.book
        # Genel memnuniyet kolonlarının yerlerini dinamik bul.
        eski_col = genel.columns.get_loc(f"Genel Memnuniyet {DONEM_1}")
        yeni_col = genel.columns.get_loc(f"Genel Memnuniyet {DONEM_2}")
        chart = karsilastirma_grafigi(
            wb, "02_Bolum_Genel_Analiz", 2, 1 + len(genel), 0, eski_col, yeni_col,
            f"{birim} Genel Memnuniyet Karşılaştırması",
        )
        ws_genel.insert_chart(2, len(genel.columns) + 2, chart, {"x_scale": 1.35, "y_scale": 1.2})

        aciklama = pd.DataFrame({
            "Başlık": [
                "Birim Kullanımı", "Karşılaştırma", "Kategori Ortalaması",
                "Genel Memnuniyet", "Özel 0 Puan", "Ağırlıklı Bölüm Ortalaması",
            ],
            "Açıklama": [
                "Kaynak dosyadaki Ders Birim değeri kullanılır. Birim aktarımı yapılmaz.",
                f"{DONEM_1} ve {DONEM_2} sonuçları aynı ders kodu + grup numarası üzerinden yan yana getirilir.",
                "MD_detay.py içindeki QUESTION_GROUPS listesinde bulunan soruların tüm geçerli cevapları birlikte ortalanır.",
                "1_1 ile 16_1 arasındaki bütün geçerli cevapların aritmetik ortalamasıdır.",
                "6_1 sorusundaki 'yapılmadı' ve 8_1 sorusundaki 'kaynak önerilmedi' cevapları hocanın referans koduna uygun olarak 0 puandır.",
                "Bölüm genelinde ders ortalamalarının basit ortalaması alınmaz; her geçerli öğrenci cevabı eşit ağırlıkta değerlendirilir. Bu nedenle katılımcısı fazla olan ders doğal olarak daha fazla katkı verir.",
            ],
        })
        dataframe_yaz(writer, "03_Hesaplama_Aciklama", aciklama,
                       f"{birim} — Hesaplama Açıklamaları")
        writer.sheets["03_Hesaplama_Aciklama"].set_column("B:B", 100)

        kullanilan = set(writer.sheets)
        anahtarlar = sorted(
            set(zip(d1["Ders Kodu"], d1["Grup No"])) | set(zip(d2["Ders Kodu"], d2["Grup No"])),
            key=lambda x: (temiz_metin(x[0]), grup_no_temizle(x[1])),
        )
        for kod, grup in anahtarlar:
            a1 = d1[(d1["Ders Kodu"] == kod) & (d1["Grup No"] == grup)].copy()
            a2 = d2[(d2["Ders Kodu"] == kod) & (d2["Grup No"] == grup)].copy()
            # attrs filtreleme sonrasında korunmayabildiği için kodları tekrar atıyoruz.
            a1.attrs["soru_kodlari"] = soru_kodlari
            a2.attrs["soru_kodlari"] = soru_kodlari
            sayfa = guvenli_sayfa_adi(f"D_{kod}_{grup}", kullanilan)
            memnuniyet_detay_sayfasi(writer, sayfa, birim, temiz_metin(kod), grup_no_temizle(grup),
                                      a1, a2, soru_kodlari, soru_metinleri)


def basari_workbook_yaz(yol: Path, birim: str, d1: pd.DataFrame, d2: pd.DataFrame):
    with pd.ExcelWriter(yol, engine="xlsxwriter") as writer:
        ozet = basari_karsilastirma(d1, d2)
        dataframe_yaz(writer, "01_Tum_Dersler_Ozet", ozet,
                       f"{birim} — Tüm Dersler Başarı Oranı Karşılaştırması",
                       ["Güncel Başarı Oranı Fark"])

        genel = basari_birim_genel(d1, d2)
        dataframe_yaz(writer, "02_Birim_Genel_Analiz", genel,
                       f"{birim} — Başarı Oranı Genel Analizi",
                       ["Güncel Başarı Oranı Fark"])

        wb = writer.book
        ws = writer.sheets["02_Birim_Genel_Analiz"]
        eski_col = genel.columns.get_loc(f"Başarı Oranı Güncel {DONEM_1}")
        yeni_col = genel.columns.get_loc(f"Başarı Oranı Güncel {DONEM_2}")
        chart = wb.add_chart({"type": "column"})
        for ad, col, renk in [(DONEM_1, eski_col, RENKLER["mavi"]), (DONEM_2, yeni_col, RENKLER["turuncu"])]:
            chart.add_series({
                "name": ad,
                "categories": ["02_Birim_Genel_Analiz", 2, 0, 1 + len(genel), 0],
                "values": ["02_Birim_Genel_Analiz", 2, col, 1 + len(genel), col],
                "fill": {"color": renk},
                "data_labels": {"value": True, "num_format": "0.00"},
            })
        chart.set_title({"name": f"{birim} Güncel Başarı Oranı Karşılaştırması"})
        chart.set_y_axis({"min": 0, "max": 100, "major_gridlines": {"visible": False}})
        chart.set_legend({"position": "bottom"})
        ws.insert_chart(2, len(genel.columns) + 2, chart, {"x_scale": 1.4, "y_scale": 1.2})

        aciklama = pd.DataFrame({
            "Alan": [
                "Başarı Oranı", "DZ", "Devamsız Hariç Öğrenci Sayısı",
                "Başarı Oranı Güncel", "Bölüm Genel Oranı",
            ],
            "Tanım / Formül": [
                "Kaynak Excel dosyasında yer alan Başarı Oranı sütunudur.",
                "Devamsız notuyla değerlendirilen öğrenci sayısıdır.",
                "Öğrenci Sayısı - DZ",
                "Başarılı Öğrenci Sayısı / (Öğrenci Sayısı - DZ) × 100",
                "Ders oranlarının basit ortalaması değildir. Toplam başarılı öğrenci / toplam devamsız-hariç öğrenci × 100 olarak hesaplanır.",
            ],
        })
        dataframe_yaz(writer, "03_Formul_Aciklama", aciklama,
                       f"{birim} — Başarı Oranı Formülleri")
        writer.sheets["03_Formul_Aciklama"].set_column("B:B", 95)

        # Ders bazında sade karşılaştırmalı grafik sayfası.
        grafik_df = ozet[[
            "Ders Kodu", "Grup No",
            f"Ders Adı {DONEM_1}", f"Ders Adı {DONEM_2}",
            f"Başarı Oranı Güncel {DONEM_1}", f"Başarı Oranı Güncel {DONEM_2}",
        ]].copy()
        grafik_df["Ders Adı"] = (
            grafik_df[f"Ders Adı {DONEM_2}"]
            .fillna(grafik_df[f"Ders Adı {DONEM_1}"])
            .fillna("")
            .map(temiz_metin)
        )
        grafik_df["Ders"] = (
            grafik_df["Ders Kodu"].astype(str)
            + grafik_df["Ders Adı"].map(lambda x: f" - {x}" if x else "")
            + " / G" + grafik_df["Grup No"].astype(str)
        )
        grafik_df = grafik_df[["Ders", f"Başarı Oranı Güncel {DONEM_1}", f"Başarı Oranı Güncel {DONEM_2}"]]
        dataframe_yaz(writer, "04_Grafikler", grafik_df,
                       f"{birim} — Ders Bazında Güncel Başarı Oranları")
        wsg = writer.sheets["04_Grafikler"]
        ch = wb.add_chart({"type": "column"})
        for ad, col, renk in [(DONEM_1, 1, RENKLER["mavi"]), (DONEM_2, 2, RENKLER["turuncu"])]:
            ch.add_series({
                "name": ad,
                "categories": ["04_Grafikler", 2, 0, 1 + len(grafik_df), 0],
                "values": ["04_Grafikler", 2, col, 1 + len(grafik_df), col],
                "fill": {"color": renk},
            })
        ch.set_title({"name": "Ders Bazında Güncel Başarı Oranı Karşılaştırması"})
        ch.set_y_axis({"min": 0, "max": 100, "major_gridlines": {"visible": False}})
        ch.set_legend({"position": "bottom"})
        ch.set_size({"width": 1200, "height": 560})
        wsg.insert_chart("E3", ch)


def kontrol_raporu_yaz(yol: Path, girdiler: GirdiDosyalari,
                        mem1: pd.DataFrame, mem2: pd.DataFrame,
                        bas1: pd.DataFrame, bas2: pd.DataFrame):
    ozet = pd.DataFrame([
        ["Memnuniyet", DONEM_1, str(girdiler.memnuniyet_1), len(mem1), mem1["Ders Birim"].nunique(), int((mem1["Ders Birim"] == BIRIMI_BOS).sum())],
        ["Memnuniyet", DONEM_2, str(girdiler.memnuniyet_2), len(mem2), mem2["Ders Birim"].nunique(), int((mem2["Ders Birim"] == BIRIMI_BOS).sum())],
        ["Başarı", DONEM_1, str(girdiler.basari_1), len(bas1), bas1["Birim"].nunique(), int((bas1["Birim"] == BIRIMI_BOS).sum())],
        ["Başarı", DONEM_2, str(girdiler.basari_2), len(bas2), bas2["Birim"].nunique(), int((bas2["Birim"] == BIRIMI_BOS).sum())],
    ], columns=["Analiz", "Dönem", "Kaynak Dosya", "İşlenen Satır", "Birim Sayısı", "Birimi Boş Satır"])

    birimler = sorted(
        set(mem1["Ders Birim"]) | set(mem2["Ders Birim"]) | set(bas1["Birim"]) | set(bas2["Birim"])
    )
    satirlar = []
    for birim in birimler:
        satirlar.append({
            "Birim": birim,
            f"Memnuniyet Satır {DONEM_1}": int((mem1["Ders Birim"] == birim).sum()),
            f"Memnuniyet Satır {DONEM_2}": int((mem2["Ders Birim"] == birim).sum()),
            f"Başarı Ders {DONEM_1}": int((bas1["Birim"] == birim).sum()),
            f"Başarı Ders {DONEM_2}": int((bas2["Birim"] == birim).sum()),
        })
    birim_df = pd.DataFrame(satirlar)
    with pd.ExcelWriter(yol, engine="xlsxwriter") as writer:
        dataframe_yaz(writer, "01_Girdi_Kontrol", ozet, "Girdi Dosyaları ve İşleme Özeti")
        dataframe_yaz(writer, "02_Birim_Kontrol", birim_df, "Birim Bazında Veri Kontrolü")
        aciklama = pd.DataFrame({
            "Kontrol": ["Birim Aktarma", "Boş Birim", "Üst Birim Filtresi", "Dönem Eşleştirme"],
            "Uygulama": [
                "Yapılmadı. Ders Birim/Birim alanı kaynaktan doğrudan kullanıldı.",
                "Memnuniyet dosyasında boş olan birimler Birimi_Bos.xlsx dosyasına alındı.",
                f"Başarı dosyasında yalnızca '{HEDEF_UST_BIRIM}' satırları işlendi.",
                "Ders Kodu + Grup No anahtarı kullanıldı; bir dönemde olmayan ders diğer dönemde boş bırakıldı.",
            ],
        })
        dataframe_yaz(writer, "03_Yontem", aciklama, "Kontrol ve Eşleştirme Yöntemi")
        writer.sheets["03_Yontem"].set_column("B:B", 100)

        # Hocanın sonradan gönderdiği resmi güncel oran sütunu varsa,
        # yeniden hesaplanan formülle uyumunu kontrol raporunda göster.
        dogrulama_satirlari = []
        for donem, veri, kaynak in [
            (DONEM_1, bas1, girdiler.basari_1),
            (DONEM_2, bas2, girdiler.basari_2),
        ]:
            bilgi = veri.attrs.get("guncel_oran_dogrulama", {})
            dogrulama_satirlari.append({
                "Dönem": donem,
                "Kaynak Dosya": str(kaynak),
                "Resmi Güncel Sütun Var": "Evet" if bilgi.get("resmi_sutun_var") else "Hayır",
                "Karşılaştırılan Satır": int(bilgi.get("karsilastirilan_satir", 0)),
                "Sıfır Payda Satırı": int(bilgi.get("sifir_payda_satir", 0)),
                "Uyumsuz Satır": int(bilgi.get("uyusmaz_satir", 0)),
                "Formül": "Başarılı / (Öğrenci - DZ) × 100",
            })
        dogrulama_df = pd.DataFrame(dogrulama_satirlari)
        dataframe_yaz(
            writer,
            "04_Guncel_Oran_Kontrol",
            dogrulama_df,
            "Resmi Güncel Başarı Oranı Sütunu Doğrulaması",
            ["Uyumsuz Satır"],
        )
        writer.sheets["04_Guncel_Oran_Kontrol"].set_column("B:B", 65)
        writer.sheets["04_Guncel_Oran_Kontrol"].set_column("G:G", 42)


# -----------------------------------------------------------------------------
# ANA AKIŞ
# -----------------------------------------------------------------------------
def bos_memnuniyet_benzeri(ornek: pd.DataFrame) -> pd.DataFrame:
    bos = ornek.iloc[0:0].copy()
    bos.attrs["soru_kodlari"] = list(ornek.attrs.get("soru_kodlari", []))
    return bos


def bos_basari_benzeri(ornek: pd.DataFrame) -> pd.DataFrame:
    return ornek.iloc[0:0].copy()


def analiz_calistir(girdiler: GirdiDosyalari, cikti: Path, ust_birim: str = HEDEF_UST_BIRIM):
    cikti.mkdir(parents=True, exist_ok=True)
    mem_dir = cikti / "Memnuniyet"
    bas_dir = cikti / "Basari_Oranlari"
    mem_dir.mkdir(exist_ok=True)
    bas_dir.mkdir(exist_ok=True)

    print("=" * 74)
    print("TBMYO İKİ DÖNEM KARŞILAŞTIRMALI ANALİZ")
    print("=" * 74)
    print("Girdi dosyaları:")
    print(f"  Memnuniyet 1 : {girdiler.memnuniyet_1}")
    print(f"  Memnuniyet 2 : {girdiler.memnuniyet_2}")
    print(f"  Başarı 1     : {girdiler.basari_1}")
    print(f"  Başarı 2     : {girdiler.basari_2}")

    print("\n[1/4] Memnuniyet dosyaları okunuyor...")
    mem1, soru_metin1 = memnuniyet_hazirla(girdiler.memnuniyet_1)
    mem2, soru_metin2 = memnuniyet_hazirla(girdiler.memnuniyet_2)
    soru_metinleri = {**soru_metin1, **soru_metin2}
    # attrs, filtrelenen alt DataFrame'lerde kaybolabildiği için ana listeyi sakla.
    ortak_sorular = sorted(set(mem1.attrs["soru_kodlari"]) | set(mem2.attrs["soru_kodlari"]), key=soru_sirasi)

    print("[2/4] Başarı dosyaları okunuyor ve üst birim süzülüyor...")
    bas1 = basari_hazirla(girdiler.basari_1, ust_birim)
    bas2 = basari_hazirla(girdiler.basari_2, ust_birim)

    # Memnuniyet kaynaklarında Ders Adı yoktur. Aynı dönemin başarı dosyasından
    # kod+grup anahtarıyla alınır; böylece özetler ve grafik başlıkları ders adını da gösterir.
    mem1 = memnuniyete_ders_adi_ekle(mem1, bas1)
    mem2 = memnuniyete_ders_adi_ekle(mem2, bas2)

    print("[3/4] Birim bazında karşılaştırmalı Excel dosyaları oluşturuluyor...")
    mem_birimler = sorted(set(mem1["Ders Birim"]) | set(mem2["Ders Birim"]))
    for sira, birim in enumerate(mem_birimler, start=1):
        d1 = mem1[mem1["Ders Birim"] == birim].copy()
        d2 = mem2[mem2["Ders Birim"] == birim].copy()
        d1.attrs["soru_kodlari"] = ortak_sorular
        d2.attrs["soru_kodlari"] = ortak_sorular
        ad = "Birimi_Bos" if birim == BIRIMI_BOS else guvenli_dosya_adi(birim)
        hedef = mem_dir / f"{ad}.xlsx"
        memnuniyet_workbook_yaz(hedef, birim, d1, d2, soru_metinleri)
        print(f"  Memnuniyet {sira:02d}/{len(mem_birimler):02d}: {hedef.name}")

    bas_birimler = sorted(set(bas1["Birim"]) | set(bas2["Birim"]))
    for sira, birim in enumerate(bas_birimler, start=1):
        d1 = bas1[bas1["Birim"] == birim].copy()
        d2 = bas2[bas2["Birim"] == birim].copy()
        ad = "Birimi_Bos" if birim == BIRIMI_BOS else guvenli_dosya_adi(birim)
        hedef = bas_dir / f"{ad}.xlsx"
        basari_workbook_yaz(hedef, birim, d1, d2)
        print(f"  Başarı     {sira:02d}/{len(bas_birimler):02d}: {hedef.name}")

    print("[4/4] Kontrol raporu yazılıyor...")
    kontrol_raporu_yaz(cikti / "00_Kontrol_Raporu.xlsx", girdiler, mem1, mem2, bas1, bas2)

    print("\n" + "=" * 74)
    print("İŞLEM TAMAMLANDI")
    print(f"Memnuniyet dosyası sayısı : {len(mem_birimler)}")
    print(f"Başarı dosyası sayısı     : {len(bas_birimler)}")
    print(f"Sonuç klasörü             : {cikti.resolve()}")
    print("=" * 74)


def argumanlar() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="TBMYO iki dönem memnuniyet ve başarı oranı karşılaştırma programı"
    )
    parser.add_argument("--memnuniyet-1", type=Path, help=f"{DONEM_1} memnuniyet Excel'i")
    parser.add_argument("--memnuniyet-2", type=Path, help=f"{DONEM_2} memnuniyet Excel'i")
    parser.add_argument("--basari-1", type=Path, help=f"{DONEM_1} başarı oranı Excel'i")
    parser.add_argument("--basari-2", type=Path, help=f"{DONEM_2} başarı oranı Excel'i")
    parser.add_argument("--cikti", type=Path, default=None, help="Sonuç klasörü")
    parser.add_argument("--ust-birim", default=HEDEF_UST_BIRIM, help="Başarı dosyasında süzülecek üst birim")
    parser.add_argument("--donem-1-adi", default=DONEM_1, help="Birinci dönemin raporlarda görünecek adı")
    parser.add_argument("--donem-2-adi", default=DONEM_2, help="İkinci dönemin raporlarda görünecek adı")
    return parser.parse_args()


def main():
    global DONEM_1, DONEM_2
    args = argumanlar()
    DONEM_1 = temiz_metin(args.donem_1_adi) or DONEM_1
    DONEM_2 = temiz_metin(args.donem_2_adi) or DONEM_2
    proje_koku = Path(__file__).resolve().parent

    tum_yollar_verildi = all([
        args.memnuniyet_1, args.memnuniyet_2, args.basari_1, args.basari_2
    ])
    otomatik = None if tum_yollar_verildi else otomatik_dosya_bul(proje_koku)
    girdiler = GirdiDosyalari(
        memnuniyet_1=args.memnuniyet_1 or otomatik.memnuniyet_1,
        memnuniyet_2=args.memnuniyet_2 or otomatik.memnuniyet_2,
        basari_1=args.basari_1 or otomatik.basari_1,
        basari_2=args.basari_2 or otomatik.basari_2,
    )
    cikti = args.cikti or (proje_koku / "Sonuclar")
    analiz_calistir(girdiler, cikti, args.ust_birim)


if __name__ == "__main__":
    try:
        main()
    except Exception as hata:
        print("\nHATA:", hata)
        print("\nAyrıntılı kullanım için: python bitirme_projesi_analiz.py --help")
        raise
