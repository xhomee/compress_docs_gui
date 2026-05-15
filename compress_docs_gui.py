"""
Стиснення фото-ксерокопій до 0.99 МБ + збірка PDF (кожна сторінка ≤ 0.99 МБ).

Встановлення залежностей:
    pip install Pillow tkinterdnd2 pikepdf

Запуск:
    python compress_docs_gui.py
"""

import os
import shutil
import threading
from pathlib import Path
from PIL import Image
import io
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

try:
    import pikepdf
    PIKEPDF_AVAILABLE = True
except ImportError:
    PIKEPDF_AVAILABLE = False

TARGET_SIZE = int(0.99 * 1024 * 1024)
SUPPORTED_EXT = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.webp'}


def get_file_size(path):
    return os.path.getsize(path)


def format_size(bytes_):
    return f"{bytes_ / (1024 * 1024):.2f} МБ"


def prepare_rgb(img):
    if img.mode in ('RGBA', 'LA', 'P'):
        background = Image.new('RGB', img.size, (255, 255, 255))
        if img.mode == 'P':
            img = img.convert('RGBA')
        background.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
        return background
    elif img.mode != 'RGB':
        return img.convert('RGB')
    return img


def compress_to_target(img, target_size):
    """Стискає PIL Image, повертає (jpeg_bytes, note)."""
    img = prepare_rgb(img)
    scale = 1.0
    min_scale = 0.3

    while scale >= min_scale:
        if scale < 1.0:
            new_size = (int(img.width * scale), int(img.height * scale))
            current_img = img.resize(new_size, Image.LANCZOS)
        else:
            current_img = img

        for quality in range(95, 25, -5):
            buffer = io.BytesIO()
            current_img.save(buffer, format='JPEG', quality=quality, optimize=True)
            if buffer.tell() <= target_size:
                note = f"якість={quality}"
                if scale < 1.0:
                    note += f", масштаб={int(scale*100)}%"
                return buffer.getvalue(), note
        scale *= 0.9

    buffer = io.BytesIO()
    final_size = (int(img.width * min_scale), int(img.height * min_scale))
    img.resize(final_size, Image.LANCZOS).save(buffer, format='JPEG', quality=30, optimize=True)
    return buffer.getvalue(), "мінімум"


def compress_image_file(input_path, output_path):
    original_size = get_file_size(input_path)
    if original_size <= TARGET_SIZE:
        shutil.copy2(input_path, output_path)
        return original_size, original_size, "без змін"
    img = Image.open(input_path)
    data, note = compress_to_target(img, TARGET_SIZE)
    with open(output_path, 'wb') as f:
        f.write(data)
    return original_size, len(data), note


def build_pdf(image_paths, output_pdf, page_target_size, progress_cb=None):
    """Кожна сторінка стискається до page_target_size з запасом для PDF-обгортки."""
    safe_target = page_target_size - 10 * 1024  # 10 КБ запас

    jpeg_pages = []
    for i, path in enumerate(image_paths, 1):
        img = Image.open(path)
        data, note = compress_to_target(img, safe_target)
        jpeg_pages.append((path.name, data, note))
        if progress_cb:
            progress_cb(i, len(image_paths), path.name, len(data), note)

    if PIKEPDF_AVAILABLE:
        _build_pdf_pikepdf(jpeg_pages, output_pdf)
    else:
        _build_pdf_pil(jpeg_pages, output_pdf)
    return jpeg_pages


def _build_pdf_pikepdf(jpeg_pages, output_pdf):
    """Вшиває JPEG напряму, без перекодування — мінімум оверхеду."""
    pdf = pikepdf.Pdf.new()
    for name, jpeg_data, note in jpeg_pages:
        img = Image.open(io.BytesIO(jpeg_data))
        width, height = img.size

        image_obj = pikepdf.Stream(pdf, jpeg_data)
        image_obj.Type = pikepdf.Name.XObject
        image_obj.Subtype = pikepdf.Name.Image
        image_obj.Width = width
        image_obj.Height = height
        image_obj.ColorSpace = pikepdf.Name.DeviceRGB
        image_obj.BitsPerComponent = 8
        image_obj.Filter = pikepdf.Name.DCTDecode

        # Створюємо сторінку через API pikepdf і встановлюємо їй вміст
        page = pdf.add_blank_page(page_size=(width, height))
        page.Resources = pikepdf.Dictionary(
            XObject=pikepdf.Dictionary(Im0=image_obj)
        )
        page.Contents = pikepdf.Stream(
            pdf,
            f"q\n{width} 0 0 {height} 0 0 cm\n/Im0 Do\nQ\n".encode()
        )
    pdf.save(output_pdf)


def _build_pdf_pil(jpeg_pages, output_pdf):
    """Резерв через PIL — простіше, але PDF трохи більший."""
    images = [Image.open(io.BytesIO(data)) for _, data, _ in jpeg_pages]
    if not images:
        return
    images[0].save(output_pdf, "PDF", save_all=True,
                   append_images=images[1:], resolution=100.0)


class CompressorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Стиснення документів + PDF")
        self.root.geometry("750x680")
        self.root.minsize(650, 550)
        self.files = []
        self.processing = False
        self._build_ui()
        if DND_AVAILABLE:
            self.drop_area.drop_target_register(DND_FILES)
            self.drop_area.dnd_bind('<<Drop>>', self.on_drop)

    def _build_ui(self):
        drop_text = "⬇  Перетягни сюди файли або папку\n\n(або натисни кнопку нижче)"
        if not DND_AVAILABLE:
            drop_text = ("Drag & drop недоступний.\n"
                         "Встанови: pip install tkinterdnd2\n\n"
                         "Використовуй кнопки нижче.")

        self.drop_area = tk.Label(
            self.root, text=drop_text, relief="groove", borderwidth=2,
            bg="#f0f0f0", font=("Arial", 12), height=4,
        )
        self.drop_area.pack(fill="x", padx=15, pady=(15, 10))

        # Список файлів
        list_frame = tk.Frame(self.root)
        list_frame.pack(fill="both", expand=False, padx=15, pady=5)
        tk.Label(list_frame, text="Файли (порядок = порядок сторінок PDF):",
                 anchor="w").pack(fill="x")
        list_container = tk.Frame(list_frame)
        list_container.pack(fill="both", expand=True)
        scrollbar_l = tk.Scrollbar(list_container)
        scrollbar_l.pack(side="right", fill="y")
        self.file_list = tk.Listbox(list_container, height=7,
                                     yscrollcommand=scrollbar_l.set,
                                     selectmode="extended")
        self.file_list.pack(side="left", fill="both", expand=True)
        scrollbar_l.config(command=self.file_list.yview)

        # Керування списком
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(fill="x", padx=15, pady=5)
        tk.Button(btn_frame, text="📄 Файли",
                  command=self.choose_files, width=10).pack(side="left", padx=2)
        tk.Button(btn_frame, text="📁 Папка",
                  command=self.choose_folder, width=10).pack(side="left", padx=2)
        tk.Button(btn_frame, text="▲ Вгору",
                  command=lambda: self.move_selected(-1), width=8).pack(side="left", padx=2)
        tk.Button(btn_frame, text="▼ Вниз",
                  command=lambda: self.move_selected(1), width=8).pack(side="left", padx=2)
        tk.Button(btn_frame, text="✕ Видалити",
                  command=self.remove_selected, width=10).pack(side="left", padx=2)
        tk.Button(btn_frame, text="🗑 Очистити",
                  command=self.clear_files, width=10).pack(side="left", padx=2)

        # Лог
        log_frame = tk.Frame(self.root)
        log_frame.pack(fill="both", expand=True, padx=15, pady=10)
        tk.Label(log_frame, text="Лог:", anchor="w").pack(fill="x")
        scrollbar = tk.Scrollbar(log_frame)
        scrollbar.pack(side="right", fill="y")
        self.log = tk.Text(log_frame, height=8, yscrollcommand=scrollbar.set,
                           font=("Consolas", 9), wrap="word")
        self.log.pack(fill="both", expand=True)
        scrollbar.config(command=self.log.yview)

        # Прогрес + статус
        self.progress = ttk.Progressbar(self.root, mode="determinate")
        self.progress.pack(fill="x", padx=15, pady=(0, 5))
        self.status = tk.Label(self.root, text="Файлів: 0", anchor="w")
        self.status.pack(fill="x", padx=15, pady=(0, 5))

        # Кнопки дій
        action_frame = tk.Frame(self.root)
        action_frame.pack(fill="x", padx=15, pady=(0, 15))
        self.compress_btn = tk.Button(
            action_frame, text="▶ Стиснути файли",
            command=self.start_compress, bg="#4CAF50", fg="white",
            font=("Arial", 10, "bold"), height=2)
        self.compress_btn.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.pdf_btn = tk.Button(
            action_frame, text="📑 Створити PDF",
            command=self.start_pdf, bg="#2196F3", fg="white",
            font=("Arial", 10, "bold"), height=2)
        self.pdf_btn.pack(side="left", fill="x", expand=True, padx=(5, 0))

    def on_drop(self, event):
        raw = event.data
        paths, current, in_braces = [], "", False
        for char in raw:
            if char == '{':
                in_braces = True
            elif char == '}':
                in_braces = False
                if current:
                    paths.append(current); current = ""
            elif char == ' ' and not in_braces:
                if current:
                    paths.append(current); current = ""
            else:
                current += char
        if current:
            paths.append(current)
        self.add_paths(paths)

    def choose_files(self):
        paths = filedialog.askopenfilenames(
            title="Виберіть фото",
            filetypes=[("Зображення", "*.jpg *.jpeg *.png *.bmp *.tiff *.webp"),
                       ("Усі файли", "*.*")]
        )
        self.add_paths(paths)

    def choose_folder(self):
        folder = filedialog.askdirectory(title="Виберіть папку з фото")
        if folder:
            self.add_paths([folder])

    def add_paths(self, paths):
        added = 0
        for p in paths:
            path = Path(p)
            if path.is_file() and path.suffix.lower() in SUPPORTED_EXT:
                if path not in self.files:
                    self.files.append(path); added += 1
            elif path.is_dir():
                for ext in SUPPORTED_EXT:
                    for f in sorted(path.glob(f'*{ext}')):
                        if f not in self.files:
                            self.files.append(f); added += 1
                    for f in sorted(path.glob(f'*{ext.upper()}')):
                        if f not in self.files:
                            self.files.append(f); added += 1
        self.refresh_list()
        if added:
            self.write_log(f"+ Додано файлів: {added}\n")

    def refresh_list(self):
        self.file_list.delete(0, "end")
        for f in self.files:
            self.file_list.insert("end", f"{f.name}  ({format_size(get_file_size(f))})")
        self.status.config(text=f"Файлів: {len(self.files)}")

    def move_selected(self, direction):
        if self.processing:
            return
        sel = list(self.file_list.curselection())
        if not sel:
            return
        if direction == -1:
            if 0 in sel:
                return
            for idx in sel:
                self.files[idx-1], self.files[idx] = self.files[idx], self.files[idx-1]
            new_sel = [i-1 for i in sel]
        else:
            if (len(self.files)-1) in sel:
                return
            for idx in reversed(sel):
                self.files[idx+1], self.files[idx] = self.files[idx], self.files[idx+1]
            new_sel = [i+1 for i in sel]
        self.refresh_list()
        for i in new_sel:
            self.file_list.selection_set(i)

    def remove_selected(self):
        if self.processing:
            return
        for idx in sorted(self.file_list.curselection(), reverse=True):
            del self.files[idx]
        self.refresh_list()

    def clear_files(self):
        if self.processing:
            return
        self.files.clear()
        self.log.delete("1.0", "end")
        self.progress["value"] = 0
        self.refresh_list()

    def write_log(self, text):
        self.log.insert("end", text)
        self.log.see("end")
        self.root.update_idletasks()

    def set_buttons(self, enabled):
        state = "normal" if enabled else "disabled"
        self.compress_btn.config(state=state)
        self.pdf_btn.config(state=state)

    def _guard_files(self):
        if not self.files:
            messagebox.showwarning("Немає файлів", "Спочатку додай файли.")
            return False
        return True

    def start_compress(self):
        if self.processing or not self._guard_files():
            return
        self.processing = True
        self.set_buttons(False)
        threading.Thread(target=self._do_compress, daemon=True).start()

    def start_pdf(self):
        if self.processing or not self._guard_files():
            return
        output_pdf = filedialog.asksaveasfilename(
            title="Зберегти PDF як...",
            defaultextension=".pdf",
            filetypes=[("PDF файли", "*.pdf")],
            initialfile="documents.pdf"
        )
        if not output_pdf:
            return
        self.processing = True
        self.set_buttons(False)
        threading.Thread(target=self._do_pdf, args=(output_pdf,), daemon=True).start()

    def _do_compress(self):
        output_dir = self.files[0].parent / "compressed"
        output_dir.mkdir(exist_ok=True)
        self.write_log(f"\n=== Стиснення ({len(self.files)} файлів) ===\n")
        self.write_log(f"Папка: {output_dir}\n\n")
        self.progress["maximum"] = len(self.files)
        self.progress["value"] = 0
        ok, fail = 0, 0
        for i, file in enumerate(self.files, 1):
            output_path = output_dir / f"{file.stem}.jpg"
            if get_file_size(file) <= TARGET_SIZE:
                output_path = output_dir / file.name
            try:
                orig, new, note = compress_image_file(file, output_path)
                arrow = "→" if orig != new else "="
                self.write_log(f"[{i}/{len(self.files)}] {file.name}\n"
                               f"    {format_size(orig)} {arrow} {format_size(new)}  ({note})\n")
                ok += 1
            except Exception as e:
                self.write_log(f"[{i}/{len(self.files)}] ✗ {file.name}: {e}\n")
                fail += 1
            self.progress["value"] = i
        self.write_log(f"\n=== Готово: {ok} ✓, {fail} ✗ ===\n")
        self._finish(ok > 0, output_dir, "Відкрити папку з результатами?")

    def _do_pdf(self, output_pdf):
        output_pdf = Path(output_pdf)
        self.write_log(f"\n=== Створення PDF ({len(self.files)} сторінок) ===\n")
        self.write_log(f"Ліміт сторінки: ≤ {format_size(TARGET_SIZE)}\n")
        if not PIKEPDF_AVAILABLE:
            self.write_log("⚠ pikepdf не встановлений — резервний метод.\n"
                           "  Для кращого результату: pip install pikepdf\n")
        self.write_log(f"Файл: {output_pdf}\n\n")
        self.progress["maximum"] = len(self.files)
        self.progress["value"] = 0

        def cb(i, total, name, size, note):
            self.write_log(f"[{i}/{total}] {name} → {format_size(size)} ({note})\n")
            self.progress["value"] = i

        try:
            build_pdf(self.files, str(output_pdf), TARGET_SIZE, progress_cb=cb)
            self.write_log(f"\n=== PDF створено ===\n")
            self.write_log(f"Розмір PDF: {format_size(get_file_size(output_pdf))}\n")
            self.write_log(f"Сторінок: {len(self.files)}\n")
            self._finish(True, output_pdf.parent, "Відкрити папку з PDF?")
        except Exception as e:
            self.write_log(f"\n✗ Помилка: {e}\n")
            self._finish(False, None, None)

    def _finish(self, success, folder, ask_text):
        self.processing = False
        self.set_buttons(True)
        if success and folder and ask_text:
            if messagebox.askyesno("Готово", ask_text):
                self.open_folder(folder)

    def open_folder(self, path):
        import platform, subprocess
        system = platform.system()
        try:
            if system == "Windows":
                os.startfile(path)
            elif system == "Darwin":
                subprocess.run(["open", path])
            else:
                subprocess.run(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося відкрити: {e}")


def main():
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    CompressorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()