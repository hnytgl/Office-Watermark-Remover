#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Lightweight Windows GUI build without PDF support.

This wrapper reuses the same Office/WPS watermark-removal core while excluding
PyMuPDF from the packaged executable. It is intended for users who only need
Word, Excel, PowerPoint, OpenDocument and convertible legacy/WPS formats.
"""
from __future__ import annotations

import office_watermark_remover as core

# Lite edition deliberately excludes PDF so PyMuPDF can be omitted from the EXE.
core.SUPPORTED_DIRECT.discard(".pdf")
core.SUPPORTED_ALL.discard(".pdf")

import office_watermark_gui as gui

# `from ... import SUPPORTED_ALL` in the GUI points at the same set, but keep this
# explicit for clarity and future refactors.
gui.SUPPORTED_ALL.discard(".pdf")
gui.APP_TITLE = "Office / WPS 去水印工具（轻量版）"
gui.FILE_TYPES = (
    ("支持的办公文件", "*.docx *.docm *.xlsx *.xlsm *.pptx *.pptm *.odt *.ods *.odp *.doc *.xls *.ppt *.wps *.et *.dps"),
    ("所有文件", "*.*"),
)


class LiteWatermarkApp(gui.WatermarkApp):
    def __init__(self) -> None:
        super().__init__()
        self.pdf_redact.set(False)
        self._patch_lite_labels(self.root)

    def _patch_lite_labels(self, widget) -> None:
        try:
            text = widget.cget("text")
        except Exception:
            text = ""

        if text == "PDF 删除指定正文文字":
            try:
                widget.configure(text="PDF 功能仅完整版提供", state="disabled")
            except Exception:
                pass
        elif isinstance(text, str) and "支持 Word、Excel、PowerPoint、PDF、OpenDocument" in text:
            try:
                widget.configure(
                    text=text.replace("、PDF", "") + "  轻量版不包含 PDF 处理组件。"
                )
            except Exception:
                pass

        try:
            children = widget.winfo_children()
        except Exception:
            children = []
        for child in children:
            self._patch_lite_labels(child)


def main() -> int:
    return LiteWatermarkApp().run()


if __name__ == "__main__":
    raise SystemExit(main())
