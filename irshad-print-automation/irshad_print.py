"""
İrşad Admin — Excel-dəki ID-ləri "Çap" siyahısına avtomatik əlavə edir.

Axın:
  1. https://manage.irshad.az saytına giriş
  2. /print/products (Məhsullar → Çap) səhifəsinə keçid
  3. Excel-dəki hər ID "Axtar" xanasına yazılıb axtarılır
  4. Nəticə sətrinin checkbox-u işarələnir (çap üçün seçim) və yaşıl "+" basılır
  5. Növbəti ID ilə təkrar

İstifadə:
  python irshad_print.py ID.xlsx
  python irshad_print.py ID.xlsx --sheet Tv --sheet Ashagi   # yalnız bu vərəqlər
  python irshad_print.py ID.xlsx --unique                    # təkrar ID-ləri bir dəfə əlavə et
  python irshad_print.py ID.xlsx --start-from 104737         # bu ID-dən davam et
  python irshad_print.py ID.xlsx --dry-run                   # yalnız ID siyahısını göstər
  python irshad_print.py ID.xlsx --date "28.09.2026 - 30.09.2026"   # endirim müddəti
  python irshad_print.py ID.xlsx --remove                    # ID-ləri çap siyahısından sil
  python irshad_print.py ID.xlsx --reset                     # brauzerdəki köhnə siyahını sıfırla
  python irshad_print.py ID.xlsx --batch 50                  # hər 50 ID-dən sonra çap üçün dayan

Endirim müddəti verilməyibsə, proqram başlanğıcda soruşur (boş buraxsanız seçilmir).
Adi Chrome açılır (ayrıca "chrome-profile" profili ilə). Bütün ID-lər əlavə
edildikdən sonra skript bitir, brauzer isə AÇIQ QALIR — çapı orada əl ilə edin.

Giriş məlumatları IRSHAD_EMAIL / IRSHAD_PASSWORD mühit dəyişənlərindən və ya
bu qovluqdakı credentials.txt faylından (1-ci sətir e-poçt, 2-ci sətir şifrə)
oxunur; heç biri yoxdursa, proqram soruşur.
"""

import argparse
import csv
import getpass
import os
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import openpyxl
from playwright.sync_api import TimeoutError as PWTimeout
from playwright.sync_api import sync_playwright

BASE_URL = "https://manage.irshad.az"
LOGIN_URL = f"{BASE_URL}/login"
PRINT_URL = f"{BASE_URL}/print/products"
HERE = Path(__file__).resolve().parent
ID_RE = re.compile(r"\d{3,8}")


def read_ids(xlsx_path, sheets=None, unique=False):
    """Bütün vərəqlərdən (və ya seçilənlərdən) ID-ləri sütun-sütun, yuxarıdan aşağı oxuyur."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)  # formulların hesablanmış dəyəri
    ids, seen = [], set()
    for ws in wb.worksheets:
        if sheets and ws.title not in sheets:
            continue
        for col in range(1, ws.max_column + 1):
            for row in range(1, ws.max_row + 1):
                v = ws.cell(row, col).value
                if isinstance(v, float) and v.is_integer():
                    v = int(v)
                s = str(v).strip() if v is not None else ""
                if not ID_RE.fullmatch(s):
                    continue
                if unique and s in seen:
                    continue
                seen.add(s)
                ids.append((ws.title, s))
    return ids


def get_credentials():
    email = os.environ.get("IRSHAD_EMAIL")
    password = os.environ.get("IRSHAD_PASSWORD")
    cred_file = HERE / "credentials.txt"
    if (not email or not password) and cred_file.exists():
        lines = [l.strip() for l in cred_file.read_text(encoding="utf-8").splitlines() if l.strip()]
        if len(lines) >= 2:
            email, password = lines[0], lines[1]
    if not email:
        email = input("E-poçt: ").strip()
    if not password:
        password = getpass.getpass("Şifrə: ")
    return email, password


def login(page, email, password):
    page.goto(LOGIN_URL, wait_until="domcontentloaded")
    page.locator("input[type=email], input[name=email]").first.fill(email)
    page.locator("input[type=password]").first.fill(password)
    page.locator("button[type=submit], button:has-text('Giriş et')").first.click()
    page.wait_for_load_state("networkidle")
    if "/login" in page.url:
        raise SystemExit("Giriş alınmadı — e-poçt/şifrəni yoxlayın.")


DATE_RE = re.compile(r"^\s*(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})\s*$")


def normalize_date_range(text):
    """'28.09.2026-30.09.2026' → '28.09.2026 - 30.09.2026'; səhvdirsə ValueError."""
    m = DATE_RE.match(text or "")
    if not m:
        raise ValueError("Tarix formatı: GG.AA.İİİİ - GG.AA.İİİİ (məs. 28.09.2026 - 30.09.2026)")
    for d in m.groups():
        try:
            datetime.strptime(d, "%d.%m.%Y")
        except ValueError:
            raise ValueError(f"Yanlış tarix: {d}")
    return f"{m.group(1)} - {m.group(2)}"


def set_discount_period(page, date_range):
    """'Endirim müddəti' xanasına tarix aralığını yazır (daterangepicker dəstəyi ilə)."""
    start, end = [d.strip() for d in date_range.split("-")]
    box = page.locator("input[placeholder='Endirim müddəti']").first
    box.wait_for(state="visible")
    box.evaluate(
        """(el, [start, end]) => {
            const value = start + ' - ' + end;
            const $ = window.jQuery;
            const drp = $ && $(el).data('daterangepicker');
            if (drp) {
                // İstifadəçinin "Tətbiq et" basmasını təqlid et: saytın öz callback-i işləsin
                drp.show();
                drp.setStartDate(start);
                drp.setEndDate(end);
                drp.clickApply();
                return;
            }
            el.value = value;
            el.dispatchEvent(new Event('input', {bubbles: true}));
            el.dispatchEvent(new Event('change', {bubbles: true}));
            if ($) $(el).trigger('change');
        }""",
        [start, end],
    )
    page.keyboard.press("Escape")  # açıq qalan təqvim pəncərəsini bağla
    actual = box.input_value()
    return actual == date_range, actual


CDP_PORT = 9333
PROFILE_DIR = HERE / "chrome-profile"


def find_browser_exe():
    """Kompüterdəki Opera-nı tapır; yoxdursa Chrome, sonra Edge."""
    env = os.environ
    candidates = []
    for base in (env.get("LOCALAPPDATA"), env.get("ProgramFiles"), env.get("ProgramFiles(x86)")):
        if base:
            for name in ("Opera", "Opera GX"):
                candidates += [Path(base) / "Programs" / name / "opera.exe", Path(base) / name / "opera.exe",
                               Path(base) / "Programs" / name / "launcher.exe"]
    for base in (env.get("ProgramFiles"), env.get("ProgramFiles(x86)"), env.get("LOCALAPPDATA")):
        if base:
            candidates.append(Path(base) / "Google/Chrome/Application/chrome.exe")
    for base in (env.get("ProgramFiles(x86)"), env.get("ProgramFiles")):
        if base:
            candidates.append(Path(base) / "Microsoft/Edge/Application/msedge.exe")
    candidates += [Path("/usr/bin/google-chrome"), Path("/usr/bin/chromium"),
                   Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")]
    for c in candidates:
        if c.exists():
            return c
    raise SystemExit("Opera, Chrome və ya Edge tapılmadı.")


def open_real_browser(p):
    """Adi brauzeri (Opera/Chrome/Edge) ayrıca proses kimi açıb ona qoşulur.

    Skript bitəndə yalnız əlaqə kəsilir — brauzer adi brauzer kimi açıq qalır və
    çap orada əl ilə edilir. Ayrıca profil (chrome-profile qovluğu) istifadə olunur,
    girişiniz yadda qalır.
    """
    url = f"http://127.0.0.1:{CDP_PORT}"
    try:  # əvvəlki işə salmadan açıq qalıbsa, ona qoşul
        return p.chromium.connect_over_cdp(url, timeout=2000)
    except Exception:
        pass
    exe = find_browser_exe()
    print(f"Brauzer: {exe}")
    flags = 0
    if sys.platform == "win32":  # skript bağlananda brauzer bağlanmasın
        flags = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
    subprocess.Popen(
        [str(exe), f"--remote-debugging-port={CDP_PORT}", f"--user-data-dir={PROFILE_DIR}",
         "--no-first-run", "--no-default-browser-check", "--start-maximized",
         "--disable-popup-blocking",  # çap səhifəsi yeni pəncərədə açıla bilsin
         "about:blank"],
        creationflags=flags, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    deadline = time.time() + 30
    while True:
        try:
            return p.chromium.connect_over_cdp(url, timeout=2000)
        except Exception:
            if time.time() > deadline:
                raise SystemExit("Brauzerə qoşulmaq alınmadı.")
            time.sleep(0.5)


def search_input(page):
    return page.locator("input[placeholder='Axtar']").first


def search_and_add(page, product_id, timeout_ms, remove=False):
    box = search_input(page)
    box.fill("")
    box.fill(product_id)
    btn = page.locator("button:has-text('Axtar'), input[type=submit][value='Axtar']")
    if btn.count():
        btn.first.click()
    else:
        box.press("Enter")
    page.wait_for_load_state("networkidle")

    # ID-si dəqiq uyğun gələn cədvəl sətri (köhnə nəticələrlə qarışmasın deyə)
    row = page.locator("table tbody tr").filter(
        has=page.locator("td", has_text=re.compile(rf"^\s*{product_id}\s*$"))
    ).first
    try:
        row.wait_for(state="visible", timeout=timeout_ms)
    except PWTimeout:
        return "tapılmadı"

    # Sətrin checkbox-u: çap düymələri ("Yeni dizayn" və s.) YALNIZ işarələnmiş məhsulları
    # çap edir — sayt onları brauzerin localStorage["print"] siyahısında saxlayır.
    checkbox = row.locator("input.check").first
    remove_sel = ".btn-danger, button:has(i.fa-minus), a:has(i.fa-minus)"
    if remove:
        if checkbox.count() and checkbox.is_checked():
            checkbox.uncheck()
        # qırmızı "−" düyməsi  # siyahıdan çıxar: qırmızı "−" düyməsi
        remove_btn = row.locator(remove_sel).first
        if not remove_btn.count():
            return "siyahıda deyil"
        remove_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(300)
        if not page.url.startswith(PRINT_URL):
            page.goto(PRINT_URL, wait_until="networkidle")
        return "silindi"

    if not checkbox.count():
        return "checkbox tapılmadı"
    checkbox.check()  # artıq işarəlidirsə toxunmur
    selected = "seçildi"

    add_btn = row.locator(
        "a.btn-success, button.btn-success, .btn-success, button:has(i.fa-plus), a:has(i.fa-plus)"
    ).first
    if not add_btn.count():
        # Qırmızı "−" düyməsi = məhsul artıq çap siyahısındadır (siyahı serverdə saxlanılır)
        remove_btn = row.locator(remove_sel)
        return f"{selected} (+ artıq əlavə olunub)" if remove_btn.count() else f"{selected} (+ düyməsi yoxdur)"
    add_btn.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(300)
    # Əlavə etdikdən sonra səhifə başqa yerə keçibsə, geri qayıt
    if not page.url.startswith(PRINT_URL):
        page.goto(PRINT_URL, wait_until="networkidle")
    return f"{selected} + əlavə edildi"


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description="İrşad Admin çap siyahısına ID-ləri avtomatik əlavə et")
    ap.add_argument("excel", help="ID-ləri olan Excel faylı (.xlsx)")
    ap.add_argument("--sheet", action="append", help="Yalnız bu vərəq(lər) (bir neçə dəfə yazmaq olar)")
    ap.add_argument("--unique", action="store_true", help="Təkrar olunan ID-ləri bir dəfə əlavə et")
    ap.add_argument("--remove", action="store_true",
                    help="Əlavə etmək əvəzinə Excel-dəki ID-ləri çap siyahısından SİL (qırmızı − düyməsi)")
    ap.add_argument("--batch", type=int, default=0,
                    help="Hər neçə ID-dən sonra çap üçün dayansın (0 = dayanmadan)")
    ap.add_argument("--reset", action="store_true",
                    help="Başlamazdan əvvəl skript brauzerində saytın məlumatını (çap siyahısı, giriş) sil")
    ap.add_argument("--start-from", help="Bu ID-dən başla (daxil olmaqla)")
    ap.add_argument("--dry-run", action="store_true", help="Sayta girmədən yalnız ID siyahısını göstər")
    ap.add_argument("--date", help='Endirim müddəti, məs. "28.09.2026 - 30.09.2026"')
    ap.add_argument("--timeout", type=int, default=10, help="Hər axtarış üçün gözləmə (saniyə)")
    args = ap.parse_args()

    ids = read_ids(args.excel, args.sheet, args.unique)
    if args.start_from:
        idx = next((i for i, (_, pid) in enumerate(ids) if pid == args.start_from), None)
        if idx is None:
            raise SystemExit(f"{args.start_from} ID-si Excel-də tapılmadı.")
        ids = ids[idx:]
    print(f"{len(ids)} ID tapıldı.")
    if args.dry_run:
        for sheet, pid in ids:
            print(f"  [{sheet}] {pid}")
        return

    date_range = "" if args.remove else args.date
    if date_range is None:
        date_range = input("Endirim müddəti (məs. 28.09.2026 - 30.09.2026, boş = seçmə): ").strip()
    if date_range:
        try:
            date_range = normalize_date_range(date_range)
        except ValueError as e:
            raise SystemExit(str(e))

    log_path = HERE / f"netice_{datetime.now():%Y%m%d_%H%M%S}.csv"
    counts = {}

    with sync_playwright() as p, open(log_path, "w", newline="", encoding="utf-8-sig") as f:
        log = csv.writer(f)
        log.writerow(["vərəq", "ID", "nəticə"])
        browser = open_real_browser(p)
        context = browser.contexts[0]
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(30_000)
        accept_dialog = lambda d: d.accept()  # əlavə zamanı çıxan təsdiq pəncərələri
        page.on("dialog", accept_dialog)

        if args.reset:
            # Bu profildə sayta aid bütün məlumatı (cookie, localStorage, çap siyahısı) sil
            cdp = context.new_cdp_session(page)
            cdp.send("Storage.clearDataForOrigin", {"origin": BASE_URL, "storageTypes": "all"})
            cdp.detach()
            print("Saytın brauzerdəki məlumatı təmizləndi.")

        page.goto(PRINT_URL, wait_until="networkidle")
        if "/login" in page.url:  # profil əvvəldən giriş etməyibsə
            login(page, *get_credentials())
            page.goto(PRINT_URL, wait_until="networkidle")

        def prepare_for_print():
            # Axtarış filtrini təmizlə və endirim müddətini yaz
            page.goto(PRINT_URL, wait_until="networkidle")
            if date_range:
                try:
                    ok, actual = set_discount_period(page, date_range)
                    print(f"Endirim müddəti: {actual}" + ("" if ok else f"  (gözlənilən: {date_range} — əl ilə yoxlayın!)"))
                except Exception as e:
                    print(f"Endirim müddəti seçilə bilmədi ({e.__class__.__name__}) — əl ilə seçin.")

        batch = args.batch if args.batch > 0 and not args.remove else len(ids)
        for n, (sheet, pid) in enumerate(ids, 1):
            try:
                result = search_and_add(page, pid, args.timeout * 1000, args.remove)
            except Exception as e:  # səhifə ilişibsə yenidən yüklə və davam et
                result = f"xəta: {e.__class__.__name__}"
                page.goto(PRINT_URL, wait_until="networkidle")
            counts[result] = counts.get(result, 0) + 1
            log.writerow([sheet, pid, result])
            f.flush()
            print(f"[{n}/{len(ids)}] {sheet} | {pid} → {result}")

            if n % batch == 0 and n < len(ids):
                # Siyahı çox böyüməsin: hər hissədən sonra çap et və təmizlə
                prepare_for_print()
                page.remove_listener("dialog", accept_dialog)
                input(f"\n>>> {n} ID əlavə olundu. Brauzerdə çap edin, sonra \"Çap siyahısını təmizlə\" "
                      f"basın və burada Enter basın (davam: {len(ids) - n} ID)... ")
                page.on("dialog", accept_dialog)
                page.goto(PRINT_URL, wait_until="networkidle")

        if not args.remove:
            prepare_for_print()

        count = page.evaluate("() => { try { return JSON.parse(localStorage.getItem('print') || '[]').length }"
                              " catch (e) { return -1 } }")
        print(f"Çap üçün seçilmiş məhsul sayı: {count}")
        page.remove_listener("dialog", accept_dialog)  # çap zamanı mesajları siz görəsiniz
        browser.close()  # yalnız əlaqəni kəsir — brauzer açıq qalır

    print("\nHazırdır. Brauzer açıq qalır — çapı orada əl ilə edin (\"Yeni dizayn\" və s.).")
    print("\nYekun:", ", ".join(f"{k}: {v}" for k, v in counts.items()))
    print(f"Hesabat: {log_path}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nDayandırıldı.")
