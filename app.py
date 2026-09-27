# -*- coding: utf-8 -*-
"""
قرار دادن عکس کالا داخل اکسل بر اساس کد حسابداری.
"""

from __future__ import annotations

import os
import re
import sys
import tempfile
import threading
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage
from PIL import ImageOps


APP_TITLE = "قرار دادن عکس در اکسل"
APP_AUTHOR = "Developed by Reza Gholami"
CODE_COLUMN_NAMES = ("کد حسابداری", "كد حسابداري", "کدحسابداری")
PHOTO_COLUMN_NAMES = ("عکس", "عكس", "تصویر", "تصوير", "photo", "image")
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tif", ".tiff"}
PERSIAN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

# اندازه نمایش داخل سلول، متعادل است؛ خود فایل با وضوح بالاتر ذخیره می‌شود تا تار نشود.
DISPLAY_MAX_PX = 168
EMBED_MAX_PX = 480


def resource_base() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def normalize_code(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip().translate(PERSIAN_DIGITS)
    if text.endswith(".0") and text[:-2].replace("-", "", 1).isdigit():
        text = text[:-2]
    text = re.sub(r"\s+", "", text)
    return text


def code_lookup_keys(code: str) -> list[str]:
    keys = []
    seen = set()
    for item in (code, code.lstrip("0") or "0" if code else ""):
        if item and item not in seen:
            keys.append(item)
            seen.add(item)
    return keys


def header_matches(cell_value: str, names: tuple[str, ...]) -> bool:
    text = normalize_code(cell_value).replace("_", "").lower()
    if not text:
        return False
    for name in names:
        compact = normalize_code(name).replace("_", "").lower()
        if text == compact or compact in text or text in compact:
            return True
    return False


def find_header_row_and_columns(ws):
    max_scan_rows = min(ws.max_row or 1, 15)
    max_scan_cols = min(ws.max_column or 1, 40)
    for row in range(1, max_scan_rows + 1):
        code_col = None
        photo_col = None
        for col in range(1, max_scan_cols + 1):
            value = ws.cell(row, col).value
            if value is None:
                continue
            text = str(value).strip()
            if code_col is None and header_matches(text, CODE_COLUMN_NAMES):
                code_col = col
            elif photo_col is None and header_matches(text, PHOTO_COLUMN_NAMES):
                photo_col = col
        if code_col:
            return row, code_col, photo_col
    return None, None, None


def build_image_index(folder: Path) -> dict[str, Path]:
    index: dict[str, Path] = {}
    for path in folder.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in IMAGE_EXTENSIONS:
            continue
        stem = normalize_code(path.stem)
        if not stem:
            continue
        for key in code_lookup_keys(stem):
            index.setdefault(key, path)

        # پشتیبانی از اسم‌هایی مثل 12345_1.jpg یا 12345-front.png
        parts = re.split(r"[_\-\s.]+", stem)
        if parts and parts[0] != stem:
            for key in code_lookup_keys(parts[0]):
                index.setdefault(key, path)
    return index


def find_image_for_code(code: str, index: dict[str, Path]) -> Path | None:
    for key in code_lookup_keys(code):
        if key in index:
            return index[key]
    return None


def fit_size(width: int, height: int, max_px: int) -> tuple[int, int]:
    longest = max(width, height)
    if longest <= 0:
        return 1, 1
    if longest <= max_px:
        return width, height
    scale = max_px / longest
    return max(1, round(width * scale)), max(1, round(height * scale))


def pixels_to_row_height(px: int) -> float:
    return px * 0.75 + 10


def pixels_to_col_width(px: int) -> float:
    return px / 7.0 + 2


def to_rgb_image(im: PILImage.Image) -> PILImage.Image:
    im = ImageOps.exif_transpose(im)
    if im.mode in ("RGBA", "LA"):
        background = PILImage.new("RGB", im.size, (255, 255, 255))
        alpha = im.getchannel("A")
        background.paste(im.convert("RGB"), mask=alpha)
        return background
    if im.mode == "P":
        return im.convert("RGBA").convert("RGB")
    if im.mode != "RGB":
        return im.convert("RGB")
    return im


def prepare_excel_image(source: Path, cache: dict[str, tuple[Path, int, int]]) -> ExcelImage:
    cached = cache.get(str(source))
    if cached is None:
        with PILImage.open(source) as im:
            im = to_rgb_image(im)
            embed_w, embed_h = fit_size(im.width, im.height, EMBED_MAX_PX)
            if (embed_w, embed_h) != (im.width, im.height):
                im = im.resize((embed_w, embed_h), PILImage.Resampling.LANCZOS)
            handle = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
            handle.close()
            im.save(handle.name, format="JPEG", quality=95, optimize=True, subsampling=0)
            cached = (Path(handle.name), embed_w, embed_h)
            cache[str(source)] = cached

    temp_path, embed_w, embed_h = cached
    display_w, display_h = fit_size(embed_w, embed_h, DISPLAY_MAX_PX)
    excel_img = ExcelImage(str(temp_path))
    excel_img.width = display_w
    excel_img.height = display_h
    return excel_img


def process_workbook(excel_path: Path, photos_dir: Path, log) -> Path:
    if not excel_path.exists():
        raise FileNotFoundError("فایل اکسل پیدا نشد.")
    if not photos_dir.exists() or not photos_dir.is_dir():
        raise FileNotFoundError("پوشه عکس پیدا نشد.")

    log("در حال خواندن اکسل...")
    keep_vba = excel_path.suffix.lower() == ".xlsm"
    wb = load_workbook(excel_path, keep_vba=keep_vba)
    ws = wb.active

    header_row, code_col, photo_col = find_header_row_and_columns(ws)
    if not header_row or not code_col:
        raise ValueError(
            "ستون «کد حسابداری» در شیت اول پیدا نشد.\n"
            "نام ستون باید دقیقاً یا شبیه «کد حسابداری» باشد."
        )

    if photo_col is None:
        photo_col = (ws.max_column or code_col) + 1
        ws.cell(header_row, photo_col).value = "عکس"
        log("ستون «عکس» وجود نداشت؛ ساخته شد.")
    else:
        log(f"ستون عکس پیدا شد: ستون {get_column_letter(photo_col)}")

    log(f"ستون کد حسابداری: {get_column_letter(code_col)} | ردیف عنوان: {header_row}")

    log("در حال فهرست کردن عکس‌ها...")
    image_index = build_image_index(photos_dir)
    if not image_index:
        raise ValueError("داخل پوشه هیچ عکس معتبری پیدا نشد.")
    log(f"{len(set(image_index.values()))} عکس برای تطبیق آماده شد.")

    # عکس‌های قبلی همان ستون را حذف می‌کنیم تا تکراری نشود
    photo_letter = get_column_letter(photo_col)
    kept_images = []
    for img in list(ws._images):
        anchor = getattr(img, "anchor", None)
        from_cell = getattr(anchor, "_from", None)
        col0 = getattr(from_cell, "col", None)
        if col0 is not None and col0 + 1 == photo_col:
            continue
        kept_images.append(img)
    ws._images = kept_images

    temp_cache: dict[str, tuple[Path, int, int]] = {}
    matched = 0
    missing = 0
    skipped = 0
    missing_codes = []
    max_display_w = DISPLAY_MAX_PX

    last_row = ws.max_row or header_row
    total_rows = max(last_row - header_row, 1)

    try:
        for row in range(header_row + 1, last_row + 1):
            raw = ws.cell(row, code_col).value
            code = normalize_code(raw)
            if not code:
                skipped += 1
                continue

            image_path = find_image_for_code(code, image_index)
            if image_path is None:
                missing += 1
                if len(missing_codes) < 30:
                    missing_codes.append(code)
                continue

            excel_img = prepare_excel_image(image_path, temp_cache)
            excel_img.anchor = f"{photo_letter}{row}"
            ws.add_image(excel_img)
            ws.row_dimensions[row].height = pixels_to_row_height(int(excel_img.height))
            max_display_w = max(max_display_w, int(excel_img.width))
            ws.cell(row, photo_col).value = None
            matched += 1

            if matched % 25 == 0:
                log(f"تا این لحظه {matched} عکس قرار گرفت...")

            progress = (row - header_row) / total_rows
            log(f"PROGRESS:{progress:.4f}")

        ws.column_dimensions[photo_letter].width = pixels_to_col_width(max_display_w)
        suffix = excel_path.suffix if excel_path.suffix.lower() in {".xlsx", ".xlsm"} else ".xlsx"
        output_path = excel_path.with_name(f"{excel_path.stem}_با_عکس{suffix}")
        log("در حال ذخیره فایل خروجی...")
        wb.save(output_path)
    finally:
        wb.close()
        for temp_file, _, _ in temp_cache.values():
            try:
                os.remove(temp_file)
            except OSError:
                pass

    log(f"PROGRESS:1")
    log("")
    log(f"عکس‌های قرارگرفته: {matched}")
    log(f"کد بدون عکس: {missing}")
    log(f"ردیف بدون کد: {skipped}")
    if missing_codes:
        extra = " ..." if missing > len(missing_codes) else ""
        log("نمونه کدهای بدون عکس: " + "، ".join(missing_codes) + extra)
    log(f"فایل ذخیره شد:\n{output_path}")
    return output_path


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE}  ·  {APP_AUTHOR}")
        self.geometry("760x640")
        self.minsize(680, 580)
        self.configure(bg="#f4f1ea")

        self.excel_path = tk.StringVar()
        self.photos_dir = tk.StringVar()
        self.status = tk.StringVar(value="فایل اکسل و پوشه عکس را انتخاب کنید.")
        self.output_path: Path | None = None
        self.worker: threading.Thread | None = None

        self._build_style()
        self._build_ui()

    def _build_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background="#f4f1ea")
        style.configure("Card.TFrame", background="#fffdf8")
        style.configure("TLabel", background="#f4f1ea", foreground="#2c2a26", font=("Segoe UI", 10))
        style.configure("Title.TLabel", background="#f4f1ea", foreground="#1f3d2a", font=("Segoe UI", 18, "bold"))
        style.configure("Hint.TLabel", background="#fffdf8", foreground="#5b574e", font=("Segoe UI", 9))
        style.configure("Card.TLabel", background="#fffdf8", foreground="#2c2a26", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10), padding=8)
        style.configure("Accent.TButton", font=("Segoe UI", 11, "bold"), padding=10)
        style.configure("Horizontal.TProgressbar", troughcolor="#e6e0d4", background="#2f6b4f")
        style.configure("Credit.TLabel", background="#f4f1ea", foreground="#6f6a60", font=("Segoe UI", 9, "italic"))

    def _build_ui(self):
        pad = {"padx": 22, "pady": 8}

        ttk.Label(self, text="قرار دادن عکس در اکسل", style="Title.TLabel").pack(anchor="e", padx=22, pady=(18, 2))
        ttk.Label(
            self,
            text="عکس هر کالا با همان کد حسابداری داخل ستون «عکس» قرار می‌گیرد.",
        ).pack(anchor="e", padx=22, pady=(0, 10))

        card = ttk.Frame(self, style="Card.TFrame", padding=16)
        card.pack(fill="x", padx=22)

        self._path_row(card, "فایل اکسل", self.excel_path, self.pick_excel)
        self._path_row(card, "پوشه عکس‌ها", self.photos_dir, self.pick_folder)

        ttk.Label(
            card,
            text="اسم هر عکس باید همان کد حسابداری باشد؛ مثلاً 1203001.jpg یا 1203001.png",
            style="Hint.TLabel",
        ).pack(anchor="e", pady=(8, 0))

        actions = ttk.Frame(self)
        actions.pack(fill="x", **pad)
        self.run_btn = ttk.Button(actions, text="شروع و ذخیره اکسل", style="Accent.TButton", command=self.start)
        self.run_btn.pack(side="right")
        self.open_btn = ttk.Button(actions, text="باز کردن پوشه خروجی", command=self.open_output, state="disabled")
        self.open_btn.pack(side="right", padx=8)

        self.progress = ttk.Progressbar(self, mode="determinate", maximum=100)
        self.progress.pack(fill="x", padx=22, pady=(4, 4))
        ttk.Label(self, textvariable=self.status).pack(anchor="e", padx=22)

        log_frame = ttk.Frame(self)
        log_frame.pack(fill="both", expand=True, padx=22, pady=(6, 8))
        self.log_box = tk.Text(
            log_frame,
            height=14,
            wrap="word",
            bg="#fffdf8",
            fg="#2c2a26",
            relief="flat",
            font=("Consolas", 10),
            padx=10,
            pady=10,
        )
        scroll = ttk.Scrollbar(log_frame, command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=scroll.set)
        self.log_box.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.log_box.insert("end", "گزارش کار اینجا نمایش داده می‌شود.\n")
        self.log_box.configure(state="disabled")
        ttk.Label(self, text=APP_AUTHOR, style="Credit.TLabel").pack(anchor="center", pady=(0, 14))

    def _path_row(self, parent, label, variable, command):
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill="x", pady=6)
        ttk.Button(row, text="انتخاب", command=command).pack(side="right")
        entry = ttk.Entry(row, textvariable=variable)
        entry.pack(side="right", fill="x", expand=True, padx=8)
        ttk.Label(row, text=label, style="Card.TLabel", width=14).pack(side="right")

    def pick_excel(self):
        path = filedialog.askopenfilename(
            title="انتخاب فایل اکسل",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("همه فایل‌ها", "*.*")],
        )
        if path:
            self.excel_path.set(path)

    def pick_folder(self):
        path = filedialog.askdirectory(title="انتخاب پوشه عکس‌ها")
        if path:
            self.photos_dir.set(path)

    def append_log(self, text: str):
        if text.startswith("PROGRESS:"):
            try:
                value = float(text.split(":", 1)[1]) * 100
                self.progress["value"] = value
            except ValueError:
                pass
            return
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        if text and not text.startswith("فایل ذخیره"):
            self.status.set(text.splitlines()[0])

    def start(self):
        excel = self.excel_path.get().strip()
        photos = self.photos_dir.get().strip()
        if not excel:
            messagebox.showwarning(APP_TITLE, "لطفاً فایل اکسل را انتخاب کنید.")
            return
        if not photos:
            messagebox.showwarning(APP_TITLE, "لطفاً پوشه عکس‌ها را انتخاب کنید.")
            return
        if self.worker and self.worker.is_alive():
            return

        self.output_path = None
        self.open_btn.configure(state="disabled")
        self.run_btn.configure(state="disabled")
        self.progress["value"] = 0
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
        self.append_log("شروع شد...")

        def work():
            try:
                output = process_workbook(Path(excel), Path(photos), lambda msg: self.after(0, self.append_log, msg))
                self.output_path = output

                def done():
                    self.run_btn.configure(state="normal")
                    self.open_btn.configure(state="normal")
                    self.status.set("تمام شد. فایل جدید ذخیره شد.")
                    messagebox.showinfo(APP_TITLE, f"اکسل همراه عکس ذخیره شد:\n{output}")

                self.after(0, done)
            except Exception as exc:
                detail = "".join(traceback.format_exception_only(type(exc), exc)).strip()

                def fail():
                    self.run_btn.configure(state="normal")
                    self.append_log("خطا: " + detail)
                    self.status.set("خطا رخ داد.")
                    messagebox.showerror(APP_TITLE, detail)

                self.after(0, fail)

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def open_output(self):
        if not self.output_path:
            return
        folder = self.output_path.parent
        if sys.platform.startswith("win"):
            os.startfile(folder)  # type: ignore[attr-defined]
        else:
            os.system(f'xdg-open "{folder}"')


def main():
    os.chdir(resource_base())
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
