"""
İrşad Admin — Excel-dəki ID-ləri "Çap" siyahısına avtomatik əlavə edir.

Axın:
  1. https://manage.irshad.az saytına giriş
  2. /print/products (Məhsullar → Çap) səhifəsinə keçid
  3. Excel-dəki hər ID "Axtar" xanasına yazılıb axtarılır
  4. Nəticə sətrinin checkbox-u işarələnir (çap üçün seçim) və yaşıl "+" basılır
  5. Növbəti ID ilə təkrar

Ara üz (pəncərə): python irshad_gui.py   — terminal rejimi aşağıdakı kimidir.

İstifadə:
  python irshad_print.py ID.xlsx
  python irshad_print.py ID.xlsx --sheet Tv --sheet Ashagi   # yalnız bu vərəqlər
  python irshad_print.py ID.xlsx --only-green                # yalnız yaşıl rəngli xanalar
  python irshad_print.py ID.xlsx --unique                    # təkrar ID-ləri bir dəfə əlavə et
  python irshad_print.py ID.xlsx --start-from 104737         # bu ID-dən davam et
  python irshad_print.py ID.xlsx --dry-run                   # yalnız ID siyahısını göstər
  python irshad_print.py ID.xlsx --date "28.09.2026 - 30.09.2026"   # endirim müddəti
  python irshad_print.py ID.xlsx --compare ID2.xlsx          # yalnız hər iki faylda olan ID-lər
  python irshad_print.py ID.xlsx --compare ID2.xlsx --compare-green   # hər iki faylda yaşıl olanlar
  python irshad_print.py ID.xlsx --reset                     # brauzerdəki köhnə siyahını sıfırla
  python irshad_print.py ID.xlsx --batch 50                  # hər 50 ID-dən sonra çap üçün dayan

Endirim müddəti verilməyibsə, proqram başlanğıcda soruşur (boş buraxsanız seçilmir).
Opera (yoxdursa Chrome/Edge) açılır (ayrıca "chrome-profile" profili ilə). Bütün ID-lər əlavə
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
FROZEN = getattr(sys, "frozen", False)  # .exe kimi işləyir (PyInstaller)
# .exe-də fayllar (credentials.txt, hesabat, profil) exe-nin yanında saxlanılır
HERE = Path(sys.executable if FROZEN else __file__).resolve().parent
ID_RE = re.compile(r"\d{3,8}")

# Filial kilidi: build_exe.bat <e-poçt> ilə qurulan exe yalnız bu hesab(lar)la işləyir.
# branch_lock.py qurulum zamanı yaradılır (git-ə düşmür); yoxdursa kilid yoxdur.
try:
    from branch_lock import ALLOWED_EMAILS
except ImportError:
    ALLOWED_EMAILS = []
ALLOWED_EMAILS = [e.strip().lower() for e in ALLOWED_EMAILS if e.strip()]


def email_allowed(email):
    return not ALLOWED_EMAILS or (email or "").strip().lower() in ALLOWED_EMAILS


def lock_message():
    return "Bu proqram yalnız bu hesab üçündür: " + ", ".join(ALLOWED_EMAILS)


MAX_COPIES = 99  # Excel-də ID-nin yanındakı 1–99 arası rəqəm = nüsxə sayı


def _theme_colors(wb):
    """Workbook temasının rəngləri (Excel-in tema indeksi sırası ilə)."""
    import xml.etree.ElementTree as ET

    default = ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31",
               "A5A5A5", "FFC000", "5B9BD5", "70AD47"]
    try:
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        scheme = ET.fromstring(wb.loaded_theme).find(".//a:clrScheme", ns)
        vals = []
        for child in list(scheme)[:10]:  # dk1, lt1, dk2, lt2, accent1..6
            el = child[0]
            vals.append(el.get("lastClr") or el.get("val"))
        # Excel tema indeksində açıq/tünd yerləri dəyişikdir: 0=lt1, 1=dk1, 2=lt2, 3=dk2
        vals[0], vals[1], vals[2], vals[3] = vals[1], vals[0], vals[3], vals[2]
        return vals
    except Exception:
        return default


def _apply_tint(rgb, tint):
    import colorsys

    r, g, b = (int(rgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + tint) if tint < 0 else l + (1 - l) * tint
    return tuple(round(c * 255) for c in colorsys.hls_to_rgb(h, l, s))


def cell_rgb(cell, theme):
    """Xananın fon rəngi (r, g, b) və ya None."""
    from openpyxl.styles.colors import COLOR_INDEX

    fill = cell.fill
    if not fill or fill.fill_type != "solid":
        return None
    c = fill.fgColor
    rgb = None
    if c.type == "rgb" and isinstance(c.rgb, str) and len(c.rgb) >= 6:
        rgb = c.rgb[-6:]
    elif c.type == "theme" and c.theme is not None and c.theme < len(theme):
        rgb = theme[c.theme]
    elif c.type == "indexed" and c.indexed is not None and c.indexed < len(COLOR_INDEX):
        rgb = COLOR_INDEX[c.indexed][-6:]
    if not rgb or rgb in ("000000",) and c.type == "indexed":
        return None
    return _apply_tint(rgb, c.tint or 0)


def is_green(rgb):
    """Rəng yaşıl çalardadırmı (açıq yaşıldan tünd yaşıla qədər)."""
    import colorsys

    if not rgb:
        return False
    h, s, v = colorsys.rgb_to_hsv(*(x / 255 for x in rgb))
    return 70 <= h * 360 <= 170 and s >= 0.12 and v >= 0.25


def _as_int(v):
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, int):
        return v
    if isinstance(v, str) and v.strip().isdigit():
        return int(v.strip())
    return None


def read_ids(xlsx_path, sheets=None, unique=False, only_green=False):
    """Excel-dən [(vərəq, ID, nüsxə sayı), ...] siyahısı.

    ID-lər sütun-sütun, yuxarıdan aşağı oxunur. Nüsxə sayı: eyni sətirdə ID-nin
    sağındakı (növbəti ID-yə qədər) ilk 1–99 arası rəqəm; yoxdursa 1.
    only_green: yalnız ID xanası (və ya yanındakı ad xanası) yaşıl rəngli olanlar.
    """
    # Rəng lazımdırsa formatla açılır; dəyərlər üçün həmişə hesablanmış (data_only) nüsxə
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    theme = _theme_colors(wb) if only_green else None
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
                if only_green:
                    colors = [cell_rgb(ws.cell(row, col), theme)]
                    if colors[0] is None and col < ws.max_column:  # ID xanası rəngsizdirsə, ad xanası
                        colors.append(cell_rgb(ws.cell(row, col + 1), theme))
                    if not any(is_green(c) for c in colors):
                        continue
                count = 1
                for c2 in range(col + 1, min(col + 5, ws.max_column + 1)):
                    v2 = ws.cell(row, c2).value
                    if v2 is not None and ID_RE.fullmatch(str(_as_int(v2) if _as_int(v2) is not None else v2).strip()):
                        break  # növbəti ID blokuna çatdıq
                    n = _as_int(v2)
                    if n is not None and 1 <= n <= MAX_COPIES:
                        count = n
                        break
                if unique and s in seen:
                    continue
                seen.add(s)
                ids.append((ws.title, s, count))
    return ids


def list_sheets(xlsx_path, only_green=False):
    """[(vərəq adı, ID sayı), ...] — ara üzdə vərəq seçimi üçün."""
    counts = {}
    for sheet, _, _ in read_ids(xlsx_path, only_green=only_green):
        counts[sheet] = counts.get(sheet, 0) + 1
    wb = openpyxl.load_workbook(xlsx_path, read_only=True)
    return [(name, counts.get(name, 0)) for name in wb.sheetnames]


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


def wait_for_manual_login(page, log=print, should_stop=None, minutes=10):
    """credentials.txt yoxdursa: istifadəçi açılan brauzerdə özü giriş edənə qədər gözlə.

    Qaytarır: giriş formasında yazılmış e-poçt (filial kilidini yoxlamaq üçün) və ya None.
    """
    from urllib.parse import parse_qs

    typed = {"email": None}

    def on_request(r):
        # Giriş formasının göndərilməsindən e-poçtu götür (şifrə heç yerdə saxlanılmır)
        if r.method == "POST" and "/login" in r.url:
            for values in parse_qs(r.post_data or "").values():
                for v in values:
                    if "@" in v:
                        typed["email"] = v.strip()

    page.on("request", on_request)
    log("Açılan brauzerdə sayta giriş edin — proqram gözləyir…")
    deadline = time.time() + minutes * 60
    while "/login" in page.url:
        if should_stop and should_stop():
            raise SystemExit("Dayandırıldı.")
        if time.time() > deadline:
            raise SystemExit("Giriş edilmədi — vaxt bitdi.")
        page.wait_for_timeout(1000)
    page.remove_listener("request", on_request)
    log("Giriş edildi.")
    return typed["email"]


def clear_site_cookies(context, page):
    """Sayt sessiyasını sil (çıxış) — filial kilidində hər dəfə yenidən giriş tələb olunur."""
    cdp = context.new_cdp_session(page)
    cdp.send("Storage.clearDataForOrigin", {"origin": BASE_URL, "storageTypes": "cookies"})
    cdp.detach()


def common_ids(items, other_path, other_green=False):
    """items-dən yalnız ikinci Excel-də də olan ID-lər (sıra və nüsxə sayı birinci fayldan).

    other_green: ikinci faylda da yalnız yaşıl xanalar nəzərə alınır.
    """
    other = {pid for _, pid, _ in read_ids(other_path, only_green=other_green)}
    return [it for it in items if it[1] in other], len(other)


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


def open_real_browser(p, log=print):
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
    log(f"Brauzer: {exe}")
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


def set_copies(row, product_id, copies):
    """Sətrin "Nüsxə sayı" xanasını yazır (saytın localStorage-i də yenilənsin deyə change ilə)."""
    box = row.locator(f'[data-count="{product_id}"]').first
    if not box.count():
        box = row.locator("input.printCount").first
    if not box.count():
        return False
    if box.input_value() != str(copies):
        box.fill(str(copies))
        box.dispatch_event("change")
    return True


def search_and_add(page, product_id, timeout_ms, copies=1):
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
    if not checkbox.count():
        return "checkbox tapılmadı"
    # Əvvəl nüsxə sayı, sonra checkbox — sayt seçimi yadda saxlayanda sayı da götürür
    copies_ok = set_copies(row, product_id, copies)
    checkbox.check()  # artıq işarəlidirsə toxunmur
    selected = "seçildi" if copies_ok or copies == 1 else "seçildi (nüsxə xanası tapılmadı!)"

    add_btn = row.locator(
        "a.btn-success, button.btn-success, .btn-success, button:has(i.fa-plus), a:has(i.fa-plus)"
    ).first
    if not add_btn.count():
        # Qırmızı "−" düyməsi = "+" artıq əvvəl basılıb
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
    ap.add_argument("excel", nargs="?", help="ID-ləri olan Excel faylı (.xlsx)")
    ap.add_argument("--sheet", action="append", help="Yalnız bu vərəq(lər) (bir neçə dəfə yazmaq olar)")
    ap.add_argument("--only-green", action="store_true",
                    help="Yalnız Excel-də yaşıl rəngli xanalardakı ID-lər")
    ap.add_argument("--unique", action="store_true", help="Təkrar olunan ID-ləri bir dəfə əlavə et")
    ap.add_argument("--compare", metavar="EXCEL2",
                    help="İkinci Excel: yalnız hər iki faylda olan ID-lər seçilir")
    ap.add_argument("--compare-green", action="store_true",
                    help="Hər iki faylda yaşıl olan ortaq ID-lər (--only-green ilə birlikdə)")
    ap.add_argument("--batch", type=int, default=0,
                    help="Hər neçə ID-dən sonra çap üçün dayansın (0 = dayanmadan)")
    ap.add_argument("--reset", action="store_true",
                    help="Başlamazdan əvvəl skript brauzerində saytın məlumatını (çap siyahısı, giriş) sil")
    ap.add_argument("--start-from", help="Bu ID-dən başla (daxil olmaqla)")
    ap.add_argument("--dry-run", action="store_true", help="Sayta girmədən yalnız ID siyahısını göstər")
    ap.add_argument("--date", help='Endirim müddəti, məs. "28.09.2026 - 30.09.2026"')
    ap.add_argument("--timeout", type=int, default=10, help="Hər axtarış üçün gözləmə (saniyə)")
    args = ap.parse_args()

    if not args.excel:
        args.excel = input("Excel faylını bura sürükləyin və Enter basın: ").strip().strip('"')

    if args.compare_green:
        args.only_green = True
    ids = read_ids(args.excel, args.sheet, args.unique, args.only_green)
    if args.compare:
        ids, n2 = common_ids(ids, args.compare, args.compare_green)
        print(f"İkinci faylda {n2} ID var; ortaq: {len(ids)}")
    if args.start_from:
        idx = next((i for i, (_, pid, _) in enumerate(ids) if pid == args.start_from), None)
        if idx is None:
            raise SystemExit(f"{args.start_from} ID-si Excel-də tapılmadı.")
        ids = ids[idx:]
    print(f"{len(ids)} ID tapıldı.")
    if args.dry_run:
        for sheet, pid, copies in ids:
            print(f"  [{sheet}] {pid}  x{copies}")
        return

    date_range = args.date
    if date_range is None:
        date_range = input("Endirim müddəti (məs. 28.09.2026 - 30.09.2026, boş = seçmə): ").strip()
    if date_range:
        try:
            date_range = normalize_date_range(date_range)
        except ValueError as e:
            raise SystemExit(str(e))

    run_job(ids, date_range, timeout=args.timeout,
            batch=args.batch, reset=args.reset)


def run_job(ids, date_range, timeout=10, batch=0, reset=False,
            log=print, progress=None, should_stop=None, credentials=None):
    """Əsas iş: ID-ləri saytda seçir. Həm terminal (CLI), həm də ara üz (GUI) istifadə edir.

    log(str) — mesaj; progress(n, total) — irəliləyiş; should_stop() — dayandırma sorğusu;
    credentials() — (e-poçt, şifrə) və ya None qaytarır, yalnız giriş lazım olanda çağırılır;
    None olarsa, istifadəçinin brauzerdə özü giriş etməsi gözlənilir.
    """
    credentials = credentials or get_credentials
    log_path = HERE / f"netice_{datetime.now():%Y%m%d_%H%M%S}.csv"
    counts = {}

    with sync_playwright() as p, open(log_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["vərəq", "ID", "nüsxə", "nəticə"])
        browser = open_real_browser(p, log)
        context = browser.contexts[0]
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(30_000)
        accept_dialog = lambda d: d.accept()  # əlavə zamanı çıxan təsdiq pəncərələri
        page.on("dialog", accept_dialog)

        if reset:
            # Bu profildə sayta aid bütün məlumatı (cookie, localStorage, çap siyahısı) sil
            cdp = context.new_cdp_session(page)
            cdp.send("Storage.clearDataForOrigin", {"origin": BASE_URL, "storageTypes": "all"})
            cdp.detach()
            log("Saytın brauzerdəki məlumatı təmizləndi.")

        if ALLOWED_EMAILS:
            # Filial kilidi: köhnə sessiyanın hansı hesaba aid olduğunu bilmirik — yenidən giriş
            clear_site_cookies(context, page)
            log("Filial kilidi: " + ", ".join(ALLOWED_EMAILS))

        page.goto(PRINT_URL, wait_until="networkidle")
        if "/login" in page.url:  # profil əvvəldən giriş etməyibsə
            creds = credentials()
            if creds:
                if not email_allowed(creds[0]):
                    raise SystemExit(lock_message() + f" (credentials.txt: {creds[0]})")
                login(page, *creds)
            else:
                email = wait_for_manual_login(page, log, should_stop)
                if ALLOWED_EMAILS and not email_allowed(email):
                    clear_site_cookies(context, page)  # icazəsiz hesabdan çıx
                    page.goto(PRINT_URL, wait_until="domcontentloaded")
                    raise SystemExit(lock_message() + (f" (giriş edilən: {email})" if email else ""))
            page.goto(PRINT_URL, wait_until="networkidle")

        def prepare_for_print():
            # Axtarış filtrini təmizlə və endirim müddətini yaz
            page.goto(PRINT_URL, wait_until="networkidle")
            if date_range:
                try:
                    ok, actual = set_discount_period(page, date_range)
                    log(f"Endirim müddəti: {actual}" + ("" if ok else f"  (gözlənilən: {date_range} — əl ilə yoxlayın!)"))
                except Exception as e:
                    log(f"Endirim müddəti seçilə bilmədi ({e.__class__.__name__}) — əl ilə seçin.")

        batch = batch if batch > 0 else len(ids)
        for n, (sheet, pid, copies) in enumerate(ids, 1):
            if should_stop and should_stop():
                log("Dayandırıldı.")
                break
            try:
                result = search_and_add(page, pid, timeout * 1000, copies)
            except Exception as e:  # səhifə ilişibsə yenidən yüklə və davam et
                result = f"xəta: {e.__class__.__name__}"
                page.goto(PRINT_URL, wait_until="networkidle")
            counts[result] = counts.get(result, 0) + 1
            writer.writerow([sheet, pid, copies, result])
            f.flush()
            log(f"[{n}/{len(ids)}] {sheet} | {pid} x{copies} → {result}")
            if progress:
                progress(n, len(ids))

            if n % batch == 0 and n < len(ids):
                # Siyahı çox böyüməsin: hər hissədən sonra çap et və təmizlə
                prepare_for_print()
                page.remove_listener("dialog", accept_dialog)
                input(f"\n>>> {n} ID əlavə olundu. Brauzerdə çap edin, sonra \"Çap siyahısını təmizlə\" "
                      f"basın və burada Enter basın (davam: {len(ids) - n} ID)... ")
                page.on("dialog", accept_dialog)
                page.goto(PRINT_URL, wait_until="networkidle")

        prepare_for_print()

        count = page.evaluate("() => { try { return JSON.parse(localStorage.getItem('print') || '[]').length }"
                              " catch (e) { return -1 } }")
        log(f"Çap üçün seçilmiş məhsul sayı: {count}")
        page.remove_listener("dialog", accept_dialog)  # çap zamanı mesajları siz görəsiniz
        browser.close()  # yalnız əlaqəni kəsir — brauzer açıq qalır

    log("\nHazırdır. Brauzer açıq qalır — çapı orada əl ilə edin (\"Yeni dizayn\" və s.).")
    log("Yekun: " + ", ".join(f"{k}: {v}" for k, v in counts.items()))
    log(f"Hesabat: {log_path}")
    return counts


if __name__ == "__main__":
    if not FROZEN:
        try:
            main()
        except KeyboardInterrupt:
            sys.exit("\nDayandırıldı.")
    else:
        try:
            main()
        except KeyboardInterrupt:
            print("\nDayandırıldı.")
        except SystemExit as e:
            if e.code not in (None, 0):
                print(e.code)
        except Exception as e:
            print(f"\nXəta: {e}")
        try:
            input("\nBağlamaq üçün Enter basın...")  # pəncərə dərhal bağlanmasın
        except EOFError:
            pass
