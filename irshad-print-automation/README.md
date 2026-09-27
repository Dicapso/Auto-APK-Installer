# İrşad Admin — Çap siyahısına avtomatik əlavə

Excel faylındakı məhsul ID-lərini `manage.irshad.az/print/products` səhifəsində
bir-bir axtarır və hər nəticənin yaşıl **+** düyməsini basır.

## Quraşdırma (Windows)
1. [Python 3](https://www.python.org/downloads/) quraşdırın ("Add to PATH" seçin).
2. Bu qovluqda `credentials.txt` yaradın (git-ə düşmür):
   ```
   e-poct@example.com
   sifre
   ```
   (Yaratmasanız proqram özü soruşacaq.)
3. Excel faylını `run.bat`-ın üzərinə sürükləyin. İlk dəfə lazımi paketlər avtomatik yüklənir.

## Əmr sətri
```
python irshad_print.py ID.xlsx                 # bütün vərəqlər
python irshad_print.py ID.xlsx --sheet Tv      # yalnız "Tv" vərəqi
python irshad_print.py ID.xlsx --unique        # təkrar ID-ləri bir dəfə
python irshad_print.py ID.xlsx --start-from 104737   # yarımçıq qalıbsa davam et
python irshad_print.py ID.xlsx --dry-run       # yalnız ID siyahısına bax
python irshad_print.py ID.xlsx --remove        # ID-ləri çap siyahısından SİL (və ya sil.bat)
python irshad_print.py ID.xlsx --date "28.09.2026 - 30.09.2026"   # endirim müddəti
```

**Endirim müddəti:** `--date` yazılmayıbsa, proqram başlanğıcda soruşur
(boş buraxsanız toxunulmur). Bütün ID-lər əlavə olunandan sonra xanaya yazılır.

**Çap:** skript kompüterdəki adi Chrome-u (yoxdursa Edge-i) ayrıca `chrome-profile`
profili ilə açır. Bütün ID-lər əlavə olunandan sonra skript bitir, brauzer isə açıq
qalır — "Yeni dizayn" və s. ilə çapı orada əl ilə edin. Giriş həmin profildə yadda
qalır; növbəti dəfə eyni brauzer pəncərəsi istifadə olunur.
