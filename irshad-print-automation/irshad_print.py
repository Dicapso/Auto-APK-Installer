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

    email, password = get_credentials()
    log_path = HERE / f"netice_{datetime.now():%Y%m%d_%H%M%S}.csv"
    counts = {}

    with sync_playwright() as p, open(log_path, "w", newline="", encoding="utf-8-sig") as f:
        log = csv.writer(f)
        log.writerow(["vərəq", "ID", "nəticə"])
        browser = p.chromium.launch(headless=args.headless, slow_mo=50)
        page = browser.new_page()
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

        browser.close()

    print("\nYekun:", ", ".join(f"{k}: {v}" for k, v in counts.items()))
    print(f"Hesabat: {log_path}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nDayandırıldı.")
