# GitHub’a Yükleme

## Önerilen depo yapısı

```text
TBMYO-Akademik-Analiz/
├── bitirme_projesi_analiz.py
├── README.md
├── HOCAYA_SAVUNMA_NOTLARI.md
├── PROJE_GEREKSINIM_KONTROL_LISTESI.md
├── requirements.txt
├── calistir.bat
├── temiz_sonuclar.bat
├── Veriler/
└── Sonuclar/
```

## Komutlar

```bash
git init
git add .
git commit -m "TBMYO iki dönem karşılaştırmalı analiz tamamlandı"
git branch -M main
git remote add origin GITHUB_DEPO_ADRESI
git push -u origin main
```

Excel verileri kişisel/kurumsal veri içeriyorsa GitHub deposunu **private** açın. Verilerin yüklenmesi istenmiyorsa `.gitignore` dosyasına şunları ekleyin:

```gitignore
Veriler/*.xlsx
Sonuclar/*.xlsx
__pycache__/
*.pyc
```
