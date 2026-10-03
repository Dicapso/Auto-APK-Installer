"""
İrşad Çap — ara üz (pəncərə).

Excel faylı, vərəqlər, yaşıl filtr və (istəyə görə) müqayisə üçün ikinci Excel burada
seçilir; iş başlayanda irəliləyiş və nəticələr eyni pəncərədə göstərilir.
Əsas iş irshad_print.run_job() funksiyasındadır.
"""

import queue
import sys
import threading
import traceback
from tkinter import filedialog, messagebox

import customtkinter as ctk

import irshad_print as core

# Saydam (şüşə) görünüş üçün rənglər — Windows-da arxa fon bulanıq (acrylic) görünür
ACCENT = "#3b82f6"
ACCENT_HOVER = "#2563eb"
DANGER = "#ef4444"
CARD = "#16161d"
CARD_BORDER = "#2a2a36"
TEXT_MUTED = "#9ca3af"
WINDOW_BG = "#0b0b10"


def glass_supported():
    """Windows-da pywinstyles varsa, bulanıq şüşə (acrylic) effekti mümkündür."""
    if sys.platform != "win32":
        return False
    try:
        import pywinstyles  # noqa: F401
        return True
    except Exception:
        return False


def apply_glass(window, glass):
    """Şüşə effekti: acrylic qara (#000000) sahələri bulanıq arxa fona çevirir."""
    if glass:
        try:
            import pywinstyles

            pywinstyles.apply_style(window, "acrylic")
            return
        except Exception:
            pass
    try:  # digər hallarda — sadəcə yarımşəffaf pəncərə
        window.attributes("-alpha", 0.94)
    except Exception:
        pass


class Card(ctk.CTkFrame):
    """Başlıqlı, yumru küncləri olan yarımşəffaf kart."""

    def __init__(self, master, title, fg_color=CARD, **kw):
        super().__init__(master, fg_color=fg_color, border_color=CARD_BORDER, border_width=1,
                         corner_radius=14, **kw)
        ctk.CTkLabel(self, text=title, font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=TEXT_MUTED, anchor="w").pack(fill="x", padx=16, pady=(12, 6))
        self.body = ctk.CTkFrame(self, fg_color="transparent")
        self.body.pack(fill="both", expand=True, padx=16, pady=(0, 14))


class App(ctk.CTk):
    def __init__(self, excel_path=None):
        self.glass = glass_supported()
        # Şüşə rejimində fon və kartlar qara — Windows onları bulanıq şəffaf göstərir
        bg = "#000000" if self.glass else WINDOW_BG
        self.card_bg = "#000000" if self.glass else CARD
        super().__init__(fg_color=bg)
        self.title("İrşad Çap")
        self.geometry("620x820")
        self.minsize(560, 640)
        self.after(50, lambda: apply_glass(self, self.glass))

        self.events = queue.Queue()
        self.stop_flag = threading.Event()
        self.worker = None
        self.sheet_vars = {}
        self.compare_path = None
        self.compare_ids = set()

        self._build()
        if excel_path:
            self._set_excel(excel_path)
        self.after(100, self._drain_events)

    # ---------- ara üz ----------
    def _card(self, master, title):
        return Card(master, title, fg_color=self.card_bg)

    def _build(self):
        root = ctk.CTkScrollableFrame(self, fg_color="transparent")
        root.pack(fill="both", expand=True, padx=14, pady=14)

        header = ctk.CTkFrame(root, fg_color="transparent")
        header.pack(fill="x", pady=(4, 12))
        ctk.CTkLabel(header, text="İrşad Çap", font=ctk.CTkFont(size=26, weight="bold"),
                     anchor="w").pack(fill="x")
        ctk.CTkLabel(header, text="Excel-dəki məhsulları çap siyahısına avtomatik seçir",
                     text_color=TEXT_MUTED, anchor="w").pack(fill="x")

        # Excel faylı
        card = self._card(root, "EXCEL FAYLI")
        card.pack(fill="x", pady=6)
        row = ctk.CTkFrame(card.body, fg_color="transparent")
        row.pack(fill="x")
        self.excel_entry = ctk.CTkEntry(row, placeholder_text="Fayl seçilməyib", height=36)
        self.excel_entry.pack(side="left", fill="x", expand=True)
        self.excel_entry.configure(state="disabled")
        ctk.CTkButton(row, text="Seç…", width=90, height=36, fg_color=ACCENT,
                      hover_color=ACCENT_HOVER, command=self._pick_excel).pack(side="left", padx=(8, 0))
        self.excel_info = ctk.CTkLabel(card.body, text="", text_color=TEXT_MUTED, anchor="w")
        self.excel_info.pack(fill="x", pady=(6, 0))
        self.only_green = ctk.CTkSwitch(card.body, text="Yalnız yaşıl rəngli xanalar", progress_color="#22c55e",
                                        command=self._on_green)
        self.only_green.pack(anchor="w", pady=(8, 0))
        ctk.CTkLabel(card.body, text="Nüsxə sayı: Excel-də ID-nin yanındakı rəqəm (məs. 2, 3); yoxdursa 1",
                     text_color=TEXT_MUTED, anchor="w", font=ctk.CTkFont(size=11)).pack(fill="x", pady=(6, 0))

        # Vərəqlər
        card = self._card(root, "VƏRƏQLƏR")
        card.pack(fill="x", pady=6)
        self.all_sheets = ctk.CTkCheckBox(card.body, text="Hamısı", command=self._toggle_all_sheets)
        self.all_sheets.pack(anchor="w")
        self.sheet_box = ctk.CTkFrame(card.body, fg_color="transparent")
        self.sheet_box.pack(fill="x", pady=(6, 0))
        ctk.CTkLabel(self.sheet_box, text="Əvvəlcə Excel faylını seçin", text_color=TEXT_MUTED).pack(anchor="w")

        # Müqayisə üçün ikinci Excel
        card = self._card(root, "MÜQAYİSƏ (İSTƏYƏ GÖRƏ)")
        card.pack(fill="x", pady=6)
        row = ctk.CTkFrame(card.body, fg_color="transparent")
        row.pack(fill="x")
        self.compare_entry = ctk.CTkEntry(row, placeholder_text="İkinci Excel seçilməyib", height=36)
        self.compare_entry.pack(side="left", fill="x", expand=True)
        self.compare_entry.configure(state="disabled")
        ctk.CTkButton(row, text="Seç…", width=70, height=36, fg_color=ACCENT,
                      hover_color=ACCENT_HOVER, command=self._pick_compare).pack(side="left", padx=(8, 0))
        ctk.CTkButton(row, text="✕", width=36, height=36, fg_color="transparent", border_width=1,
                      border_color=CARD_BORDER, hover_color=CARD_BORDER,
                      command=self._clear_compare).pack(side="left", padx=(6, 0))
        self.compare_info = ctk.CTkLabel(card.body, text="Seçilsə, yalnız HƏR İKİ faylda olan ID-lər çapa verilir",
                                         text_color=TEXT_MUTED, anchor="w", font=ctk.CTkFont(size=11))
        self.compare_info.pack(fill="x", pady=(6, 0))

        # Başlat / Dayandır
        row = ctk.CTkFrame(root, fg_color="transparent")
        row.pack(fill="x", pady=(12, 6))
        self.start_btn = ctk.CTkButton(row, text="Başlat", height=44, fg_color=ACCENT,
                                       hover_color=ACCENT_HOVER, font=ctk.CTkFont(size=15, weight="bold"),
                                       command=self._start)
        self.start_btn.pack(side="left", fill="x", expand=True)
        self.stop_btn = ctk.CTkButton(row, text="Dayandır", width=120, height=44, fg_color="transparent",
                                      border_width=1, border_color=DANGER, text_color=DANGER,
                                      hover_color="#2a1215", state="disabled", command=self._stop)
        self.stop_btn.pack(side="left", padx=(8, 0))

        # İrəliləyiş və jurnal
        card = self._card(root, "GEDİŞAT")
        card.pack(fill="both", expand=True, pady=6)
        self.progress = ctk.CTkProgressBar(card.body, progress_color=ACCENT, height=8)
        self.progress.set(0)
        self.progress.pack(fill="x")
        self.status = ctk.CTkLabel(card.body, text="Hazır", text_color=TEXT_MUTED, anchor="w")
        self.status.pack(fill="x", pady=(6, 6))
        self.log_box = ctk.CTkTextbox(card.body, height=200, fg_color="#0f0f15",
                                      font=ctk.CTkFont(family="Consolas", size=12))
        self.log_box.pack(fill="both", expand=True)
        self.log_box.configure(state="disabled")

    # ---------- köməkçilər ----------
    def _pick_excel(self):
        path = filedialog.askopenfilename(title="Excel faylını seçin",
                                          filetypes=[("Excel", "*.xlsx *.xlsm"), ("Hamısı", "*.*")])
        if path:
            self._set_excel(path)

    def _set_excel(self, path):
        try:
            sheets = core.list_sheets(path, bool(self.only_green.get()))
        except Exception as e:
            messagebox.showerror("Xəta", f"Excel faylı oxunmadı:\n{e}")
            return
        self.excel_path = path
        self._set_entry(self.excel_entry, path)

        for w in self.sheet_box.winfo_children():
            w.destroy()
        self.sheet_vars = {}
        self._sheet_counts = dict(sheets)
        for name, count in sheets:
            var = ctk.BooleanVar(value=count > 0)
            cb = ctk.CTkCheckBox(self.sheet_box, text=f"{name}   ({count} ID)", variable=var,
                                 command=self._refresh_count)
            if count == 0:
                cb.configure(state="disabled")
            cb.pack(anchor="w", pady=2)
            self.sheet_vars[name] = var
        self.all_sheets.select()
        self._refresh_count()

    def _on_green(self):
        # Yaşıl filtr dəyişəndə vərəqlərdəki ID sayları yenidən hesablanır
        if getattr(self, "excel_path", None):
            self._set_excel(self.excel_path)

    def _toggle_all_sheets(self):
        on = bool(self.all_sheets.get())
        for name, var in self.sheet_vars.items():
            var.set(on and self._sheet_counts.get(name, 0) > 0)
        self._refresh_count()

    def _pick_compare(self):
        path = filedialog.askopenfilename(title="Müqayisə üçün ikinci Excel",
                                          filetypes=[("Excel", "*.xlsx *.xlsm"), ("Hamısı", "*.*")])
        if not path:
            return
        try:
            ids = {pid for _, pid, _ in core.read_ids(path)}
        except Exception as e:
            messagebox.showerror("Xəta", f"Excel faylı oxunmadı:\n{e}")
            return
        self.compare_path, self.compare_ids = path, ids
        self._set_entry(self.compare_entry, path)
        self._refresh_count()

    def _clear_compare(self):
        self.compare_path, self.compare_ids = None, set()
        self._set_entry(self.compare_entry, "")
        self.compare_info.configure(text="Seçilsə, yalnız HƏR İKİ faylda olan ID-lər çapa verilir")
        self._refresh_count()

    @staticmethod
    def _set_entry(entry, text):
        entry.configure(state="normal")
        entry.delete(0, "end")
        entry.insert(0, text)
        entry.configure(state="disabled")

    def _first_file_ids(self):
        sheets = [n for n, v in self.sheet_vars.items() if v.get()]
        if not sheets:
            return []
        # Eyni ID bir dəfə seçilir (nüsxə sayı ilk rast gəlinən sətirdən)
        return core.read_ids(self.excel_path, sheets, True, bool(self.only_green.get()))

    def _selected_ids(self):
        items = self._first_file_ids()
        if self.compare_path:
            items = [it for it in items if it[1] in self.compare_ids]
        return items

    def _refresh_count(self):
        if not getattr(self, "excel_path", None):
            return
        first = self._first_file_ids()
        items = self._selected_ids()
        n, copies = len(items), sum(c for _, _, c in items)
        if self.compare_path:
            self.compare_info.configure(
                text=f"İkinci faylda {len(self.compare_ids)} ID · 1-ci faylda {len(first)} · ortaq: {n}")
        chosen = sum(v.get() for v in self.sheet_vars.values())
        enabled = sum(1 for name, c in self._sheet_counts.items() if c > 0)
        (self.all_sheets.select if chosen == enabled else self.all_sheets.deselect)()
        self.excel_info.configure(text=f"{n} ID seçilib · {copies} nüsxə · {chosen}/{len(self.sheet_vars)} vərəq")

    def _log(self, text):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _set_running(self, running):
        self.start_btn.configure(state="disabled" if running else "normal",
                                 text="İşləyir…" if running else "Başlat")
        self.stop_btn.configure(state="normal" if running else "disabled")

    # ---------- iş ----------
    def _start(self):
        if not getattr(self, "excel_path", None):
            messagebox.showwarning("Excel", "Əvvəlcə Excel faylını seçin.")
            return
        ids = self._selected_ids()
        if not ids:
            if self.compare_path and self._first_file_ids():
                messagebox.showwarning("Müqayisə", "İki faylda ortaq ID yoxdur.")
            else:
                messagebox.showwarning("Vərəqlər", "Ən azı bir vərəq seçin.")
            return

        def credentials():
            # credentials.txt varsa avtomatik giriş; yoxdursa istifadəçi brauzerdə özü giriş edir
            cred = core.HERE / "credentials.txt"
            if cred.exists():
                lines = [l.strip() for l in cred.read_text(encoding="utf-8").splitlines() if l.strip()]
                if len(lines) >= 2:
                    return lines[0], lines[1]
            return None

        self.stop_flag.clear()
        self.progress.set(0)
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
        self._set_running(True)
        self.status.configure(text=f"Başladı: {len(ids)} ID")

        def work():
            try:
                core.run_job(
                    ids, "",
                    log=lambda m: self.events.put(("log", m)),
                    progress=lambda n, t: self.events.put(("progress", (n, t))),
                    should_stop=self.stop_flag.is_set,
                    credentials=credentials,
                )
                self.events.put(("done", None))
            except SystemExit as e:
                self.events.put(("error", str(e.code)))
            except Exception as e:
                self.events.put(("log", traceback.format_exc()))
                self.events.put(("error", f"{e.__class__.__name__}: {e}"))

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def _stop(self):
        self.stop_flag.set()
        self.status.configure(text="Dayandırılır… (cari ID bitəndən sonra)")

    def _drain_events(self):
        try:
            while True:
                kind, data = self.events.get_nowait()
                if kind == "log":
                    self._log(data)
                elif kind == "progress":
                    n, total = data
                    self.progress.set(n / total)
                    self.status.configure(text=f"{n} / {total}")
                elif kind == "done":
                    self._set_running(False)
                    self.status.configure(text="Hazırdır — brauzerdə çapı edin")
                    self.progress.set(1)
                elif kind == "error":
                    self._set_running(False)
                    self.status.configure(text="Xəta")
                    self._log(f"Xəta: {data}")
                    messagebox.showerror("Xəta", data)
        except queue.Empty:
            pass
        self.after(100, self._drain_events)


def main():
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    excel = sys.argv[1] if len(sys.argv) > 1 else None  # exe-nin üzərinə Excel sürüklənibsə
    App(excel).mainloop()


if __name__ == "__main__":
    main()
