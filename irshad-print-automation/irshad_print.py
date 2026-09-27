"""
İrşad Admin — Excel-dəki ID-ləri "Çap" siyahısına avtomatik əlavə edir.

Axın:
  1. https://manage.irshad.az saytına giriş
  2. /print/products (Məhsullar → Çap) səhifəsinə keçid
  3. Excel-dəki hər ID "Axtar" xanasına yazılıb axtarılır
  4. Nəticə sətrindəki yaşıl "+" düyməsi basılır
  5. Növbəti ID ilə təkrar

İstifadə:
  python irshad_print.py ID.xlsx
  python irshad_print.py ID.xlsx --sheet Tv --sheet Ashagi   # yalnız bu vərəqlər
  python irshad_print.py ID.xlsx --unique                    # təkrar ID-ləri bir dəfə əlavə et
  python irshad_print.py ID.xlsx --start-from 104737         # bu ID-dən davam et
  python irshad_print.py ID.xlsx --dry-run                   # yalnız ID siyahısını göstər
  python irshad_print.py ID.xlsx --date "28.09.2026 - 30.09.2026"   # endirim müddəti

Endirim müddəti verilməyibsə, proqram başlanğıcda soruşur (boş buraxsanız seçilmir).
Bütün ID-lər əlavə edildikdən sonra brauzer AÇIQ QALIR — çapı əl ilə edin,
bitirəndə brauzer pəncərəsini bağlayın.

Giriş məlumatları IRSHAD_EMAIL / IRSHAD_PASSWORD mühit dəyişənlərindən və ya
bu qovluqdakı credentials.txt faylından (1-ci sətir e-poçt, 2-ci sətir şifrə)
oxunur; heç biri yoxdursa, proqram soruşur.
"""

import argparse
import csv
import getpass
import os
import re
import sys
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


def watch_page(pg, download_dir):
    """Çap zamanı brauzerin gizli davranışlarını istifadəçiyə görünən et.

    Playwright susmaya görə alert/confirm pəncərələrini dərhal bağlayır və
    yüklənən faylları (məs. çap PDF-i) gizli qovluğa atır — ona görə "heç nə olmur".
    """
    def on_dialog(d):
        print(f"  [sayt mesajı] {d.message}")
        d.accept()

    def on_download(dl):
        target = download_dir / dl.suggested_filename
        dl.save_as(target)
        print(f"  [fayl yükləndi] {target}")
        if hasattr(os, "startfile"):
            os.startfile(target)  # Windows-da PDF-i standart proqramla aç

    pg.on("dialog", on_dialog)
    pg.on("download", on_download)


def launch_browser(p, headless):
    """Kompüterdəki real Chrome/Edge-i açır (PDF baxıcısı və çap normal işləsin)."""
    for channel in ("chrome", "msedge", None):
        try:
            browser = p.chromium.launch(channel=channel, headless=headless, slow_mo=50)
            print(f"Brauzer: {channel or 'chromium (daxili)'}")
            return browser
        except Exception:
            continue
    raise SystemExit("Brauzer açıla bilmədi.")


def search_input(page):
    return page.locator("input[placeholder='Axtar']").first


def search_and_add(page, product_id, timeout_ms):
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

    add_btn = row.locator(
        "a.btn-success, button.btn-success, .btn-success, button:has(i.fa-plus), a:has(i.fa-plus)"
    ).first
    if not add_btn.count():
        return "düymə tapılmadı"
    add_btn.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(300)
    # Əlavə etdikdən sonra səhifə başqa yerə keçibsə, geri qayıt
    if not page.url.startswith(PRINT_URL):
        page.goto(PRINT_URL, wait_until="networkidle")
    return "əlavə edildi"


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    ap = argparse.ArgumentParser(description="İrşad Admin çap siyahısına ID-ləri avtomatik əlavə et")
    ap.add_argument("excel", help="ID-ləri olan Excel faylı (.xlsx)")
    ap.add_argument("--sheet", action="append", help="Yalnız bu vərəq(lər) (bir neçə dəfə yazmaq olar)")
    ap.add_argument("--unique", action="store_true", help="Təkrar olunan ID-ləri bir dəfə əlavə et")
    ap.add_argument("--start-from", help="Bu ID-dən başla (daxil olmaqla)")
    ap.add_argument("--dry-run", action="store_true", help="Sayta girmədən yalnız ID siyahısını göstər")
    ap.add_argument("--headless", action="store_true", help="Brauzer pəncərəsini göstərmə")
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

    date_range = args.date
    if date_range is None:
        date_range = input("Endirim müddəti (məs. 28.09.2026 - 30.09.2026, boş = seçmə): ").strip()
    if date_range:
        try:
            date_range = normalize_date_range(date_range)
        except ValueError as e:
            raise SystemExit(str(e))

    email, password = get_credentials()
    log_path = HERE / f"netice_{datetime.now():%Y%m%d_%H%M%S}.csv"
    counts = {}

    with sync_playwright() as p, open(log_path, "w", newline="", encoding="utf-8-sig") as f:
        log = csv.writer(f)
        log.writerow(["vərəq", "ID", "nəticə"])
        browser = launch_browser(p, args.headless)
        context = browser.new_context(accept_downloads=True, no_viewport=True)
        download_dir = HERE / "cap"
        download_dir.mkdir(exist_ok=True)
        # Çap düyməsi yeni tab açsa, orada da mesaj/yükləmələri tut
        context.on("page", lambda pg: watch_page(pg, download_dir))
        page = context.new_page()
        page.set_default_timeout(30_000)

        login(page, email, password)
        page.goto(PRINT_URL, wait_until="networkidle")

        for n, (sheet, pid) in enumerate(ids, 1):
            try:
                result = search_and_add(page, pid, args.timeout * 1000)
            except Exception as e:  # səhifə ilişibsə yenidən yüklə və davam et
                result = f"xəta: {e.__class__.__name__}"
                page.goto(PRINT_URL, wait_until="networkidle")
            counts[result] = counts.get(result, 0) + 1
            log.writerow([sheet, pid, result])
            f.flush()
            print(f"[{n}/{len(ids)}] {sheet} | {pid} → {result}")

        if date_range:
            try:
                ok, actual = set_discount_period(page, date_range)
                print(f"Endirim müddəti: {actual}" + ("" if ok else f"  (gözlənilən: {date_range} — əl ilə yoxlayın!)"))
            except Exception as e:
                print(f"Endirim müddəti seçilə bilmədi ({e.__class__.__name__}) — əl ilə seçin.")

        print("\nHazırdır. Brauzer açıq qalır — çapı əl ilə edin.")
        print("Bitirəndə brauzer pəncərəsini bağlayın (proqram özü bağlanacaq).")
        # Bütün pəncərələr bağlanana qədər gözlə (bu vaxt mesaj/yükləmə hadisələri işlənir)
        while browser.is_connected():
            open_pages = [pg for pg in context.pages if not pg.is_closed()]
            if not open_pages:
                break
            try:
                open_pages[0].wait_for_timeout(500)
            except Exception:
                pass
        try:
            browser.close()
        except Exception:
            pass

    print("\nYekun:", ", ".join(f"{k}: {v}" for k, v in counts.items()))
    print(f"Hesabat: {log_path}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nDayandırıldı.")
