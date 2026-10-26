"""Offline desktop interface for the DOCX formatter (macOS and Windows)."""
from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from formatter import FormatError, format_docx, inspect, load_templates


BG = "#F7F5F0"
WHITE = "#FFFFFF"
INK = "#1C2733"
MUTED = "#718087"
TEAL = "#126F69"
LINE = "#DEE6DF"
WARNING = "#946237"


def human_size(byte_count: int) -> str:
    return f"{byte_count / 1024:.1f} KB" if byte_count < 1024 * 1024 else f"{byte_count / 1024 / 1024:.1f} MB"


class DesktopApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("公文格式助手")
        self.geometry("820x690")
        self.minsize(680, 590)
        self.configure(bg=BG)
        self.selected_path: Path | None = None
        self.last_output: Path | None = None
        self.busy = False
        self.templates = load_templates()
        self.template_ids = [item["id"] for item in self.templates]
        self.font_name = "Microsoft YaHei" if sys.platform == "win32" else "PingFang SC"
        self._build()

    def _label(self, parent, text, size=12, color=INK, bold=False, **kwargs):
        return tk.Label(parent, text=text, font=(self.font_name, size, "bold" if bold else "normal"),
                        fg=color, bg=parent["bg"], **kwargs)

    def _build(self) -> None:
        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=46, pady=(34, 24))
        self._label(header, "文", 20, WHITE, True, width=2, height=1).configure(bg=TEAL)
        header.winfo_children()[-1].pack(side="left")
        self._label(header, "公文格式助手", 17, INK, True).pack(side="left", padx=13)
        self._label(header, "离线桌面版", 10, MUTED).pack(side="right")

        intro = tk.Frame(self, bg=BG)
        intro.pack(fill="x", padx=46, pady=(0, 25))
        self._label(intro, "格式有标准，内容更从容。", 25, INK, True).pack(anchor="w")
        self._label(intro, "选择 Word 文档，检查识别结果，再保存调整后的文件。文档只在这台电脑上处理。", 11, MUTED).pack(anchor="w", pady=(9, 0))

        card = tk.Frame(self, bg=WHITE, highlightbackground=LINE, highlightthickness=1)
        card.pack(fill="both", expand=True, padx=46, pady=(0, 20))
        content = tk.Frame(card, bg=WHITE)
        content.pack(fill="both", expand=True, padx=28, pady=24)

        self._label(content, "01  选择格式模板", 10, TEAL, True).pack(anchor="w")
        self.template_name = tk.StringVar(value=self.templates[0]["name"])
        style = ttk.Style(self)
        style.configure("Format.TCombobox", padding=8, font=(self.font_name, 11))
        self.template_box = ttk.Combobox(content, textvariable=self.template_name,
                                         values=[item["name"] for item in self.templates], state="readonly", style="Format.TCombobox")
        self.template_box.pack(fill="x", pady=(9, 21))
        self.template_box.bind("<<ComboboxSelected>>", lambda _: self._reanalyze())

        self._label(content, "02  选择 Word 文档", 10, TEAL, True).pack(anchor="w")
        upload = tk.Frame(content, bg="#F7FBF8", highlightbackground="#B8D2C2", highlightthickness=1)
        upload.pack(fill="x", pady=(9, 12), ipady=17)
        self.file_name = self._label(upload, "尚未选择文档", 12, INK, True)
        self.file_name.pack(pady=(4, 6), padx=16)
        self.file_info = self._label(upload, "支持 .docx · 最大 15 MB", 10, MUTED)
        self.file_info.pack(pady=(0, 12))
        tk.Button(upload, text="选择文档", command=self.choose_file,
                  font=(self.font_name, 11, "bold"), fg=TEAL, bg=WHITE, relief="flat",
                  padx=22, pady=8, cursor="hand2").pack()

        self._label(content, "03  检查与保存", 10, TEAL, True).pack(anchor="w", pady=(8, 8))
        report_frame = tk.Frame(content, bg="#FBFCFA", highlightbackground=LINE, highlightthickness=1)
        report_frame.pack(fill="both", expand=True)
        self.report_text = tk.Text(report_frame, wrap="word", height=7, relief="flat", bg="#FBFCFA",
                                   fg=INK, font=(self.font_name, 10), padx=14, pady=12, spacing1=3, spacing3=6)
        scrollbar = ttk.Scrollbar(report_frame, orient="vertical", command=self.report_text.yview)
        self.report_text.configure(yscrollcommand=scrollbar.set)
        self.report_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._set_report("请选择文档以查看识别结果。")

        actions = tk.Frame(content, bg=WHITE)
        actions.pack(fill="x", pady=(18, 0))
        self.save_button = tk.Button(actions, text="保存调整后的文档  →", command=self.save_formatted,
                                     font=(self.font_name, 11, "bold"), fg=WHITE, bg=TEAL,
                                     activebackground="#075E59", activeforeground=WHITE,
                                     relief="flat", padx=22, pady=11, state="disabled", cursor="hand2")
        self.save_button.pack(side="left")
        self.open_button = tk.Button(actions, text="打开输出文件", command=self.open_output,
                                     font=(self.font_name, 10), fg=TEAL, bg=WHITE,
                                     relief="flat", padx=16, pady=10, state="disabled", cursor="hand2")
        self.open_button.pack(side="left", padx=12)
        self.status = self._label(actions, "就绪", 10, MUTED)
        self.status.pack(side="right")

        footer = tk.Frame(self, bg=BG)
        footer.pack(fill="x", padx=46, pady=(0, 22))
        self._label(footer, "依据 GB/T 9704—2012 调整常见排版；印章、红头、页码等请人工复核。", 10, MUTED).pack(anchor="w")

    def _set_report(self, message: str) -> None:
        self.report_text.configure(state="normal")
        self.report_text.delete("1.0", "end")
        self.report_text.insert("1.0", message)
        self.report_text.configure(state="disabled")

    def _template_id(self) -> str:
        index = [item["name"] for item in self.templates].index(self.template_name.get())
        return self.template_ids[index]

    def choose_file(self) -> None:
        path = filedialog.askopenfilename(title="选择 Word 文档", filetypes=[("Word 文档", "*.docx")])
        if not path:
            return
        candidate = Path(path)
        if candidate.suffix.lower() != ".docx":
            messagebox.showerror("文件格式", "目前只支持 .docx Word 文档。")
            return
        self.selected_path = candidate
        self.last_output = None
        self.file_name.configure(text=candidate.name)
        self.file_info.configure(text=human_size(candidate.stat().st_size))
        self.open_button.configure(state="disabled")
        self._reanalyze()

    def _reanalyze(self) -> None:
        if not self.selected_path or self.busy:
            return
        self.busy = True
        self.save_button.configure(state="disabled")
        self.status.configure(text="正在检查…")
        path, template = self.selected_path, self._template_id()

        def work():
            try:
                report = inspect(path.read_bytes(), template)
                self.after(0, lambda: self._analysis_done(report))
            except (OSError, FormatError) as exc:
                self.after(0, lambda error=str(exc): self._failed(error))

        threading.Thread(target=work, daemon=True).start()

    def _analysis_done(self, report) -> None:
        self.busy = False
        lines = [f"识别标题：{report['title'] or '未检测到'}",
                 f"段落数量：{report['paragraphs']}    处理依据：{report['basis']}"]
        if report["warnings"]:
            lines.extend(["", "需要人工复核："] + [f"• {item}" for item in report["warnings"]])
        self._set_report("\n".join(lines))
        self.save_button.configure(state="normal")
        self.status.configure(text="检查完成")

    def _failed(self, error: str) -> None:
        self.busy = False
        self.save_button.configure(state="disabled")
        self.status.configure(text="处理失败")
        self._set_report(error)
        messagebox.showerror("无法处理文档", error)

    def save_formatted(self) -> None:
        if not self.selected_path or self.busy:
            return
        output = filedialog.asksaveasfilename(title="保存调整后的文档", defaultextension=".docx",
                     initialfile=f"{self.selected_path.stem}_标准格式.docx",
                     filetypes=[("Word 文档", "*.docx")])
        if not output:
            return
        output_path = Path(output)
        if output_path.resolve() == self.selected_path.resolve():
            messagebox.showerror("保存位置", "请选择不同于原文档的文件名，避免覆盖原件。")
            return
        self.busy = True
        self.save_button.configure(state="disabled")
        self.status.configure(text="正在排版…")
        source, template = self.selected_path, self._template_id()

        def work():
            try:
                result, report = format_docx(source.read_bytes(), template)
                temp_path = None
                try:
                    with tempfile.NamedTemporaryFile(prefix="gongwen-", suffix=".docx", dir=output_path.parent, delete=False) as stream:
                        temp_path = Path(stream.name)
                        stream.write(result)
                    os.replace(temp_path, output_path)
                finally:
                    if temp_path and temp_path.exists():
                        temp_path.unlink()
                self.after(0, lambda: self._save_done(output_path, report))
            except (OSError, FormatError) as exc:
                self.after(0, lambda error=str(exc): self._failed(error))

        threading.Thread(target=work, daemon=True).start()

    def _save_done(self, path: Path, report) -> None:
        self.busy = False
        self.last_output = path
        self.save_button.configure(state="normal")
        self.open_button.configure(state="normal")
        self.status.configure(text="保存完成")
        messagebox.showinfo("排版完成", f"已保存到：\n{path}\n\n正式使用前，请复核文档中的红头、印章、页码和复杂版式。")

    def open_output(self) -> None:
        if not self.last_output or not self.last_output.exists():
            return
        try:
            if sys.platform == "win32":
                os.startfile(str(self.last_output))
            else:
                subprocess.Popen(["open", str(self.last_output)])
        except OSError as exc:
            messagebox.showerror("无法打开文档", str(exc))


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        templates = load_templates()
        assert templates and templates[0]["id"] == "gov_official"
        print("desktop bundle OK")
    else:
        DesktopApp().mainloop()
