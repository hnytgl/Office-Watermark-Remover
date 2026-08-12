#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Windows full GUI with PDF -> Word -> watermark removal -> PDF as default."""
from __future__ import annotations

from pathlib import Path

import office_watermark_gui as gui
from pdf_roundtrip import process_pdf_roundtrip

_original_process_one = gui.process_one


def _smart_process_one(
    src,
    output,
    texts,
    regex,
    auto,
    aggressive,
    pdf_redact_text,
    soffice,
    dry_run=False,
):
    path = Path(src)
    if path.suffix.lower() == ".pdf" and not dry_run:
        return process_pdf_roundtrip(
            path,
            Path(output) if output else None,
            texts,
            regex,
            auto,
            aggressive,
            fallback_direct=True,
        )
    return _original_process_one(
        src,
        output,
        texts,
        regex,
        auto,
        aggressive,
        pdf_redact_text,
        soffice,
        dry_run,
    )


# Existing worker resolves process_one from the GUI module at runtime, so replace
# it here without duplicating the entire GUI implementation.
gui.process_one = _smart_process_one
gui.APP_TITLE = "Office / WPS / PDF 去水印工具（PDF推荐模式）"


class RoundtripWatermarkApp(gui.WatermarkApp):
    def __init__(self) -> None:
        super().__init__()
        self._log(
            "PDF 默认流程：PDF → Microsoft Word → Word 去水印 → PDF；"
            "需要 Windows 已安装桌面版 Microsoft Word。转换失败时自动回退到原 PDF 直接处理。"
        )
        self._patch_pdf_labels(self.root)

    def _patch_pdf_labels(self, widget) -> None:
        try:
            text = widget.cget("text")
        except Exception:
            text = ""
        if text == "PDF 删除指定正文文字":
            try:
                widget.configure(
                    text="PDF 水印文字（可选）：推荐模式会自动识别；此项用于回退直接模式"
                )
            except Exception:
                pass
        try:
            children = widget.winfo_children()
        except Exception:
            children = []
        for child in children:
            self._patch_pdf_labels(child)


def main() -> int:
    return RoundtripWatermarkApp().run()


if __name__ == "__main__":
    raise SystemExit(main())
