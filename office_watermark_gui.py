#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from office_watermark_remover import LEGACY_CONVERT, SUPPORTED_ALL, process_one

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:
    DND_FILES = None
    TkinterDnD = None

APP_TITLE = "Office / WPS / PDF 去水印工具"
FILE_TYPES = (
    ("支持的办公文件", "*.docx *.docm *.xlsx *.xlsm *.pptx *.pptm *.pdf *.odt *.ods *.odp *.doc *.xls *.ppt *.wps *.et *.dps"),
    ("所有文件", "*.*"),
)


class WatermarkApp:
    def __init__(self) -> None:
        root_cls = TkinterDnD.Tk if TkinterDnD else tk.Tk
        self.root = root_cls()
        self.root.title(APP_TITLE)
        self.root.geometry("900x680")
        self.root.minsize(760, 560)

        self.files: list[Path] = []
        self.output_dir = tk.StringVar()
        self.watermark_text = tk.StringVar()
        self.auto_mode = tk.BooleanVar(value=True)
        self.aggressive = tk.BooleanVar(value=False)
        self.pdf_redact = tk.BooleanVar(value=False)
        self.dry_run = tk.BooleanVar(value=False)
        self.progress = tk.DoubleVar(value=0)
        self.status = tk.StringVar(value="请选择文件，或把文件拖到窗口中。")
        self._worker: threading.Thread | None = None
        self._run_options: dict[str, object] = {}
        self._events: queue.Queue[tuple[str, object]] = queue.Queue()

        self._build_ui()
        self._setup_drop()
        self.root.after(120, self._poll_events)

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}
        outer = ttk.Frame(self.root)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text=APP_TITLE, font=("Microsoft YaHei UI", 16, "bold")).pack(
            anchor="w", padx=14, pady=(14, 4)
        )
        ttk.Label(
            outer,
            text="支持 Word、Excel、PowerPoint、PDF、OpenDocument，以及部分 WPS/旧 Office 格式。默认采用安全识别，不覆盖原文件。",
            wraplength=850,
        ).pack(anchor="w", padx=14, pady=(0, 8))

        toolbar = ttk.Frame(outer)
        toolbar.pack(fill="x", **pad)
        ttk.Button(toolbar, text="添加文件", command=self.add_files).pack(side="left", padx=(0, 6))
        ttk.Button(toolbar, text="添加文件夹", command=self.add_folder).pack(side="left", padx=6)
        ttk.Button(toolbar, text="移除选中", command=self.remove_selected).pack(side="left", padx=6)
        ttk.Button(toolbar, text="清空", command=self.clear_files).pack(side="left", padx=6)

        list_frame = ttk.Frame(outer)
        list_frame.pack(fill="both", expand=True, padx=14, pady=(0, 8))
        self.tree = ttk.Treeview(
            list_frame, columns=("type", "status"), show="tree headings", selectmode="extended"
        )
        self.tree.heading("#0", text="文件")
        self.tree.heading("type", text="格式")
        self.tree.heading("status", text="状态")
        self.tree.column("#0", width=590, minwidth=300)
        self.tree.column("type", width=80, anchor="center")
        self.tree.column("status", width=160, anchor="center")
        ybar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=ybar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        ybar.pack(side="right", fill="y")

        options = ttk.LabelFrame(outer, text="处理选项")
        options.pack(fill="x", padx=14, pady=6)

        row1 = ttk.Frame(options)
        row1.pack(fill="x", **pad)
        ttk.Label(row1, text="指定水印文字：").pack(side="left")
        ttk.Entry(row1, textvariable=self.watermark_text).pack(side="left", fill="x", expand=True, padx=(6, 8))
        ttk.Label(row1, text="可留空；多个关键词用 | 分隔").pack(side="left")

        row2 = ttk.Frame(options)
        row2.pack(fill="x", **pad)
        ttk.Checkbutton(row2, text="安全自动识别", variable=self.auto_mode).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(row2, text="激进模式（可能误删普通图形）", variable=self.aggressive).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(row2, text="PDF 删除指定正文文字", variable=self.pdf_redact).pack(side="left", padx=(0, 14))
        ttk.Checkbutton(row2, text="只扫描不输出", variable=self.dry_run).pack(side="left")

        row3 = ttk.Frame(options)
        row3.pack(fill="x", **pad)
        ttk.Label(row3, text="输出目录：").pack(side="left")
        ttk.Entry(row3, textvariable=self.output_dir).pack(side="left", fill="x", expand=True, padx=(6, 8))
        ttk.Button(row3, text="选择", command=self.choose_output_dir).pack(side="left", padx=(0, 6))
        ttk.Button(row3, text="清除", command=lambda: self.output_dir.set("")).pack(side="left")
        ttk.Label(
            options,
            text="留空时在原文件旁生成“_无水印”文件；始终保留原文件。",
            foreground="#666",
        ).pack(anchor="w", padx=10, pady=(0, 8))

        actions = ttk.Frame(outer)
        actions.pack(fill="x", padx=14, pady=(6, 4))
        self.start_btn = ttk.Button(actions, text="开始处理", command=self.start_processing)
        self.start_btn.pack(side="left")
        ttk.Button(actions, text="打开输出目录", command=self.open_output_dir).pack(side="left", padx=8)
        self.progressbar = ttk.Progressbar(actions, variable=self.progress, maximum=100)
        self.progressbar.pack(side="left", fill="x", expand=True, padx=(12, 0))

        ttk.Label(outer, textvariable=self.status).pack(anchor="w", padx=14, pady=(0, 4))

        log_frame = ttk.LabelFrame(outer, text="处理日志")
        log_frame.pack(fill="both", expand=False, padx=14, pady=(0, 14))
        self.log = tk.Text(log_frame, height=9, wrap="word", state="disabled")
        logbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=logbar.set)
        self.log.pack(side="left", fill="both", expand=True)
        logbar.pack(side="right", fill="y")

    def _setup_drop(self) -> None:
        if not DND_FILES:
            self._log("提示：未安装 tkinterdnd2，仍可用“添加文件/文件夹”。安装后可直接拖拽文件。")
            return
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind("<<Drop>>", self._on_drop)
        self._log("已启用拖拽：可把文件或文件夹直接拖到文件列表。")

    def _on_drop(self, event) -> None:
        try:
            items = list(self.root.tk.splitlist(event.data))
        except Exception:
            items = [event.data]
        self.add_paths([Path(x) for x in items])

    def _log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text.rstrip() + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def add_files(self) -> None:
        files = filedialog.askopenfilenames(title="选择需要去水印的文件", filetypes=FILE_TYPES)
        self.add_paths([Path(x) for x in files])

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(title="选择包含办公文件的文件夹")
        if folder:
            self.add_paths([Path(folder)])

    def add_paths(self, paths: list[Path]) -> None:
        known = {str(p.resolve()).lower() for p in self.files if p.exists()}
        added = 0
        for path in paths:
            if path.is_dir():
                candidates = [p for p in path.rglob("*") if p.is_file()]
            else:
                candidates = [path]
            for p in candidates:
                if p.suffix.lower() not in SUPPORTED_ALL or "_无水印" in p.stem:
                    continue
                try:
                    key = str(p.resolve()).lower()
                except Exception:
                    key = str(p).lower()
                if key in known:
                    continue
                known.add(key)
                self.files.append(p)
                self.tree.insert(
                    "", "end", iid=str(len(self.files) - 1), text=str(p), values=(p.suffix.lower(), "待处理")
                )
                added += 1
        self.status.set(f"已添加 {len(self.files)} 个文件。")
        if added == 0 and paths:
            self._log("没有新增可处理的文件。")

    def remove_selected(self) -> None:
        selected_paths = {self.tree.item(i, "text") for i in self.tree.selection()}
        if not selected_paths:
            return
        self.files = [p for p in self.files if str(p) not in selected_paths]
        self._refresh_tree()
        self.status.set(f"已保留 {len(self.files)} 个文件。")

    def clear_files(self) -> None:
        self.files.clear()
        self._refresh_tree()
        self.progress.set(0)
        self.status.set("文件列表已清空。")

    def _refresh_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, p in enumerate(self.files):
            self.tree.insert("", "end", iid=str(idx), text=str(p), values=(p.suffix.lower(), "待处理"))

    def choose_output_dir(self) -> None:
        folder = filedialog.askdirectory(title="选择输出目录")
        if folder:
            self.output_dir.set(folder)

    def _texts(self) -> list[str]:
        return [x.strip() for x in self.watermark_text.get().split("|") if x.strip()]

    def start_processing(self) -> None:
        if self._worker and self._worker.is_alive():
            return
        if not self.files:
            messagebox.showwarning(APP_TITLE, "请先添加需要处理的文件。")
            return
        if self.pdf_redact.get() and not self._texts():
            messagebox.showwarning(APP_TITLE, "启用“PDF 删除指定正文文字”时，必须填写水印文字。")
            return
        if self.aggressive.get():
            ok = messagebox.askyesno(
                APP_TITLE, "激进模式可能误删普通页眉、图形或背景。\n\n确认继续吗？"
            )
            if not ok:
                return

        self._run_options = {
            "texts": self._texts(),
            "auto": bool(self.auto_mode.get()),
            "aggressive": bool(self.aggressive.get()),
            "pdf_redact": bool(self.pdf_redact.get()),
            "dry_run": bool(self.dry_run.get()),
            "output_dir": self.output_dir.get().strip(),
        }
        self.start_btn.configure(state="disabled")
        self.progress.set(0)
        self.status.set("正在处理……")
        self._log("=" * 64)
        self._worker = threading.Thread(target=self._process_worker, daemon=True)
        self._worker.start()

    def _process_worker(self) -> None:
        opts = dict(self._run_options)
        texts = list(opts.get("texts", []))
        auto = bool(opts.get("auto", True))
        aggressive = bool(opts.get("aggressive", False))
        pdf_redact = bool(opts.get("pdf_redact", False))
        dry_run = bool(opts.get("dry_run", False))
        output_dir = str(opts.get("output_dir", ""))
        total = len(self.files)
        success = 0
        changed = 0
        failures = 0
        outputs: list[str] = []

        for index, src in enumerate(list(self.files), 1):
            self._events.put(("item", (index - 1, "处理中")))
            output = None
            if output_dir:
                ext = LEGACY_CONVERT.get(src.suffix.lower(), src.suffix)
                output = Path(output_dir) / f"{src.stem}_无水印{ext}"
            result = process_one(src, output, texts, False, auto, aggressive, pdf_redact, None, dry_run)
            if result.error:
                failures += 1
                status = "失败"
                self._events.put(("log", f"[失败] {src}\n  {result.error}"))
            else:
                success += 1
                changed += result.count
                status = f"完成 / {result.count}项"
                if result.output:
                    outputs.append(result.output)
                detail = f"[完成] {src}：删除/修改 {result.count} 个水印相关对象"
                if result.output:
                    detail += f"\n  输出：{result.output}"
                for warning in result.warnings:
                    detail += f"\n  提示：{warning}"
                self._events.put(("log", detail))
            self._events.put(("item", (index - 1, status)))
            self._events.put(("progress", index * 100.0 / total))

        self._events.put(("done", (success, failures, changed, outputs)))

    def _poll_events(self) -> None:
        try:
            while True:
                kind, payload = self._events.get_nowait()
                if kind == "log":
                    self._log(str(payload))
                elif kind == "progress":
                    self.progress.set(float(payload))
                elif kind == "item":
                    idx, status = payload
                    iid = str(idx)
                    if self.tree.exists(iid):
                        vals = list(self.tree.item(iid, "values"))
                        while len(vals) < 2:
                            vals.append("")
                        vals[1] = status
                        self.tree.item(iid, values=vals)
                elif kind == "done":
                    success, failures, changed, outputs = payload
                    self.start_btn.configure(state="normal")
                    self.status.set(
                        f"处理完成：成功 {success}，失败 {failures}，共删除/修改 {changed} 个水印相关对象。"
                    )
                    msg = self.status.get()
                    if outputs:
                        msg += "\n\n输出文件已生成。"
                    messagebox.showinfo(APP_TITLE, msg)
        except queue.Empty:
            pass
        self.root.after(120, self._poll_events)

    def open_output_dir(self) -> None:
        folder = self.output_dir.get().strip()
        if not folder:
            if not self.files:
                messagebox.showinfo(APP_TITLE, "尚未选择文件或输出目录。")
                return
            folder = str(self.files[0].parent)
        path = Path(folder)
        if not path.exists():
            messagebox.showwarning(APP_TITLE, "输出目录不存在。")
            return
        try:
            if os.name == "nt":
                os.startfile(path)
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}" >/dev/null 2>&1 &')
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"无法打开目录：{exc}")

    def run(self) -> int:
        self.root.mainloop()
        return 0


def main() -> int:
    return WatermarkApp().run()


if __name__ == "__main__":
    raise SystemExit(main())
