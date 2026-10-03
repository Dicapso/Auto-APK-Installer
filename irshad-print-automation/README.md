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
3. `run.bat`-ı açın (və ya Excel faylını onun üzərinə sürükləyin) — **ara üz** açılır:
   Excel faylı, vərəqlər, yaşıl filtr, müqayisə üçün ikinci Excel (istəyə görə) və
   **Başlat** düyməsi. Gedişat və nəticələr eyni pəncərədə görünür.
   İlk dəfə lazımi paketlər avtomatik yüklənir. Windows 10/11-də pəncərə şəffaf (acrylic) görünür.

## Əmr sətri (ara üzsüz)
```
python irshad_print.py ID.xlsx                 # bütün vərəqlər
python irshad_print.py ID.xlsx --sheet Tv      # yalnız "Tv" vərəqi
python irshad_print.py ID.xlsx --only-green    # yalnız yaşıl rəngli xanalar
python irshad_print.py ID.xlsx --unique        # təkrar ID-ləri bir dəfə
python irshad_print.py ID.xlsx --start-from 104737   # yarımçıq qalıbsa davam et
python irshad_print.py ID.xlsx --dry-run       # yalnız ID siyahısına bax
python irshad_print.py ID.xlsx --compare ID2.xlsx   # yalnız hər iki faylda olan ID-lər
python irshad_print.py ID.xlsx --reset         # skript brauzerindəki köhnə siyahını sıfırla 
python irshad_print.py ID.xlsx --batch 50      # hər 50 ID-dən sonra çap üçün dayan 
python irshad_print.py ID.xlsx --date "28.09.2026 - 30.09.2026"   # endirim müddəti
```

**Endirim müddəti:** `--date` yazılmayıbsa, proqram başlanğıcda soruşur
(boş buraxsanız toxunulmur). Bütün ID-lər əlavə olunandan sonra xanaya yazılır.

**Çap:** "Yeni dizayn" və digər çap düymələri yalnız **checkbox-u işarələnmiş** məhsulları
çap edir (sayt onları brauzerin yaddaşında saxlayır). Skript hər ID üçün sətrin checkbox-unu
işarələyir və yaşıl **+** düyməsini basır. Brauzer olaraq Opera (yoxdursa Chrome, sonra Edge)
ayrıca `chrome-profile` profili ilə açılır; skript bitəndə brauzer açıq qalır — çapı orada edin.

## Kodu gizlətmək: .exe
`build_exe.bat`-ı işə salın — `dist\IrshadCap.exe` (ara üzlü) yaranır. Başqalarına yalnız bu exe
faylını verin (Python lazım deyil). Exe-ni iki dəfə klikləyin və ya Excel faylını onun
üzərinə sürükləyin. `credentials.txt`, `chrome-profile` və hesabat faylları exe-nin
yanında yaranır — bunları paylaşmayın.

**Müqayisə:** ikinci Excel seçilsə, yalnız hər iki faylda olan ID-lər çapa verilir
(sıra, vərəq/yaşıl filtr və nüsxə sayı birinci fayldan). **Giriş:** `credentials.txt`
varsa avtomatik; yoxdursa açılan brauzerdə özünüz giriş edirsiniz, proqram gözləyir.
Endirim müddətini saytda özünüz seçirsiniz.
