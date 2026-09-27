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
```

ID-lər bütün vərəqlərdən sütun-sütun, yuxarıdan aşağı oxunur (3–8 rəqəmli dəyərlər;
formul xanalarının hesablanmış dəyəri götürülür). Hər ID-nin nəticəsi
`netice_<tarix>.csv` faylına yazılır (`əlavə edildi` / `tapılmadı` / `xəta`).
