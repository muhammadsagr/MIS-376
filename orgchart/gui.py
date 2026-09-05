"""واجهة رسومية بسيطة (Tkinter) لمن يفضّل الاستخدام بالفأرة."""

from __future__ import annotations

import os
import threading
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .cli import _parse_slides
from .converter import convert, default_output_path


class ConverterApp:
    def __init__(self, root: tk.Tk, initial_file: Optional[str] = None):
        self.root = root
        root.title("تحويل الهيكل الوظيفي من PowerPoint إلى Excel")
        root.geometry("760x520")
        root.minsize(680, 480)

        self.pptx_var = tk.StringVar(value=initial_file or "")
        self.xlsx_var = tk.StringVar()
        self.slides_var = tk.StringVar()
        self.titles_var = tk.BooleanVar(value=False)
        self.tables_var = tk.BooleanVar(value=True)
        self.smart_var = tk.BooleanVar(value=True)
        self.geom_var = tk.BooleanVar(value=True)
        self.rtl_var = tk.BooleanVar(value=True)
        self.positions_var = tk.BooleanVar(value=True)

        self._build()
        if initial_file:
            self.xlsx_var.set(default_output_path(initial_file))

    # ---------------------------------------------------------------- layout
    def _build(self) -> None:
        pad = {"padx": 8, "pady": 6}
        frame = ttk.LabelFrame(self.root, text="الملفات")
        frame.pack(fill="x", **pad)

        ttk.Label(frame, text="ملف PowerPoint:").grid(row=0, column=0, sticky="e", **pad)
        ttk.Entry(frame, textvariable=self.pptx_var, width=60).grid(row=0, column=1, **pad)
        ttk.Button(frame, text="استعراض...", command=self.pick_input).grid(row=0, column=2, **pad)

        ttk.Label(frame, text="ملف Excel الناتج:").grid(row=1, column=0, sticky="e", **pad)
        ttk.Entry(frame, textvariable=self.xlsx_var, width=60).grid(row=1, column=1, **pad)
        ttk.Button(frame, text="حفظ باسم...", command=self.pick_output).grid(row=1, column=2, **pad)

        options = ttk.LabelFrame(self.root, text="الخيارات")
        options.pack(fill="x", **pad)
        ttk.Label(options, text="الشرائح (مثال 1,3,5-7):").grid(row=0, column=0, sticky="e", **pad)
        ttk.Entry(options, textvariable=self.slides_var, width=20).grid(row=0, column=1, sticky="w", **pad)
        ttk.Checkbutton(options, text="قراءة الجداول", variable=self.tables_var)\
            .grid(row=0, column=2, sticky="w", **pad)
        ttk.Checkbutton(options, text="إدراج عناوين الشرائح", variable=self.titles_var)\
            .grid(row=1, column=0, sticky="w", **pad)
        ttk.Checkbutton(options, text="فصل ذكي للاسم والمسمى", variable=self.smart_var)\
            .grid(row=1, column=1, sticky="w", **pad)
        ttk.Checkbutton(options, text="استنتاج التسلسل من المواقع", variable=self.geom_var)\
            .grid(row=1, column=2, sticky="w", **pad)
        ttk.Checkbutton(options, text="اتجاه الإكسل من اليمين لليسار", variable=self.rtl_var)\
            .grid(row=2, column=0, sticky="w", **pad)
        ttk.Checkbutton(options, text="هيكل وظائف (وليس موظفين)", variable=self.positions_var)\
            .grid(row=2, column=1, sticky="w", **pad)

        actions = ttk.Frame(self.root)
        actions.pack(fill="x", **pad)
        self.convert_btn = ttk.Button(actions, text="تحويل الآن", command=self.start_convert)
        self.convert_btn.pack(side="right", padx=8)
        ttk.Button(actions, text="فتح مجلد الناتج", command=self.open_folder).pack(side="right")

        log_frame = ttk.LabelFrame(self.root, text="النتيجة")
        log_frame.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(log_frame, height=12, wrap="word")
        self.log.pack(fill="both", expand=True, padx=6, pady=6)

    # ---------------------------------------------------------------- actions
    def write(self, message: str) -> None:
        self.log.insert("end", message + "\n")
        self.log.see("end")
        self.root.update_idletasks()

    def pick_input(self) -> None:
        path = filedialog.askopenfilename(
            title="اختر ملف العرض التقديمي",
            filetypes=[("PowerPoint", "*.pptx *.pptm"), ("All files", "*.*")])
        if path:
            self.pptx_var.set(path)
            self.xlsx_var.set(default_output_path(path))

    def pick_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="احفظ ملف الإكسل", defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if path:
            self.xlsx_var.set(path)

    def open_folder(self) -> None:
        path = self.xlsx_var.get() or self.pptx_var.get()
        folder = os.path.dirname(os.path.abspath(path)) if path else os.getcwd()
        try:
            if os.name == "nt":
                os.startfile(folder)                                # noqa: S606
            else:
                import subprocess
                subprocess.Popen(["xdg-open", folder])
        except Exception as exc:
            messagebox.showwarning("تنبيه", f"تعذّر فتح المجلد: {exc}")

    def start_convert(self) -> None:
        self.convert_btn.state(["disabled"])
        threading.Thread(target=self._convert, daemon=True).start()

    def _convert(self) -> None:
        try:
            source = self.pptx_var.get().strip()
            if not source:
                messagebox.showerror("خطأ", "الرجاء اختيار ملف PowerPoint أولاً.")
                return
            self.write(f"جاري قراءة: {source}")
            output, nodes, result = convert(
                pptx_path=source,
                excel_path=self.xlsx_var.get().strip() or None,
                slides=_parse_slides(self.slides_var.get().strip()),
                include_titles=self.titles_var.get(),
                use_tables=self.tables_var.get(),
                smart_text=self.smart_var.get(),
                infer_geometry=self.geom_var.get(),
                rtl=self.rtl_var.get(),
                positions=self.positions_var.get(),
            )
            self.xlsx_var.set(output)
            self.write(f"عدد الوظائف المستخرجة: {len(nodes)}")
            self.write(f"عدد المستويات: {max((n.level for n in nodes), default=0)}")
            for warning in result.warnings:
                self.write(f"تنبيه: {warning}")
            self.write(f"تم الحفظ في: {output}")
            messagebox.showinfo("تم بنجاح", f"تم إنشاء الملف:\n{output}")
        except Exception as exc:
            self.write(f"خطأ: {exc}")
            messagebox.showerror("خطأ", str(exc))
        finally:
            self.convert_btn.state(["!disabled"])


def run_gui(initial_file: Optional[str] = None) -> int:
    root = tk.Tk()
    try:
        ttk.Style().theme_use("clam")
    except Exception:
        pass
    ConverterApp(root, initial_file)
    root.mainloop()
    return 0
