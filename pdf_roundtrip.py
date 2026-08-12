#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PDF watermark removal through a Microsoft Word round-trip.

Flow: PDF -> Word DOCX -> remove watermark objects/text -> PDF.
This keeps the existing Office watermark-removal core and avoids bundling a
second heavy PDF-to-DOCX engine. It requires Microsoft Word on Windows.
"""
from __future__ import annotations

import math
import os
import re
import shutil
import subprocess
import tempfile
from collections import Counter
from pathlib import Path
from typing import Sequence

import office_watermark_remover as core


def _norm(s: str) -> str:
    return re.sub(r"\s+", "", s or "").lower()


def discover_pdf_watermark_texts(src: Path, max_candidates: int = 8) -> list[str]:
    """Find likely repeated/diagonal text watermarks before conversion.

    This is intentionally conservative. A diagonal short line appearing on
    several pages is a strong watermark signal. Known watermark words are also
    accepted. For a one-page PDF, diagonal text is accepted without repetition.
    """
    try:
        import pymupdf
    except Exception:
        return []

    doc = pymupdf.open(src)
    try:
        page_count = max(1, doc.page_count)
        counts: Counter[str] = Counter()
        original: dict[str, str] = {}
        diagonal_count: Counter[str] = Counter()

        for page in doc:
            seen: set[str] = set()
            seen_diagonal: set[str] = set()
            try:
                data = page.get_text("dict")
            except Exception:
                continue
            for block in data.get("blocks", []):
                for line in block.get("lines", []):
                    spans = line.get("spans", [])
                    text = "".join(str(s.get("text", "")) for s in spans).strip()
                    text = re.sub(r"\s+", " ", text)
                    if len(text) < 2 or len(text) > 100:
                        continue
                    key = _norm(text)
                    if not key:
                        continue
                    direction = line.get("dir", (1.0, 0.0))
                    try:
                        dx, dy = float(direction[0]), float(direction[1])
                    except Exception:
                        dx, dy = 1.0, 0.0
                    diagonal = abs(dy) >= 0.12 or abs(dx) <= 0.97
                    known = core.looks_like_watermark_text(text)
                    if diagonal or known:
                        original.setdefault(key, text)
                        seen.add(key)
                        if diagonal:
                            seen_diagonal.add(key)
            counts.update(seen)
            diagonal_count.update(seen_diagonal)

        threshold = 1 if page_count == 1 else max(2, math.ceil(page_count * 0.4))
        ranked: list[tuple[int, int, int, str]] = []
        for key, count in counts.items():
            text = original[key]
            known = 1 if core.looks_like_watermark_text(text) else 0
            diagonal = diagonal_count[key]
            if known or count >= threshold or (page_count == 1 and diagonal):
                ranked.append((known, count, diagonal, text))
        ranked.sort(reverse=True)
        return [item[3] for item in ranked[:max_candidates]]
    finally:
        doc.close()


def _powershell_executable() -> str | None:
    for name in ("powershell.exe", "powershell", "pwsh.exe", "pwsh"):
        path = shutil.which(name)
        if path:
            return path
    return None


def _run_word_com(input_file: Path, output_file: Path, operation: str) -> None:
    if os.name != "nt":
        raise RuntimeError("PDF→Word→PDF 模式目前仅支持 Windows，并需要安装 Microsoft Word。")
    ps = _powershell_executable()
    if not ps:
        raise RuntimeError("未找到 PowerShell，无法调用 Microsoft Word 转换。")

    if operation == "pdf_to_docx":
        action = r'''
$word.Options.ConfirmConversions = $false
$doc = $word.Documents.Open($InputFile, $false, $true, $false)
$doc.SaveAs2($OutputFile, 12)
'''
    elif operation == "docx_to_pdf":
        action = r'''
$doc = $word.Documents.Open($InputFile, $false, $true, $false)
$doc.ExportAsFixedFormat($OutputFile, 17)
'''
    else:
        raise ValueError(operation)

    script = rf'''
param([string]$InputFile, [string]$OutputFile)
$ErrorActionPreference = "Stop"
$word = $null
$doc = $null
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    {action}
    if ($doc -ne $null) {{ $doc.Close(0) }}
    $doc = $null
}} finally {{
    if ($doc -ne $null) {{ try {{ $doc.Close(0) }} catch {{}} }}
    if ($word -ne $null) {{ try {{ $word.Quit() }} catch {{}} }}
}}
'''
    with tempfile.TemporaryDirectory(prefix="wm_wordcom_") as td:
        ps1 = Path(td) / "convert.ps1"
        ps1.write_text(script, encoding="utf-8-sig")
        cmd = [
            ps,
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(ps1),
            "-InputFile",
            str(input_file.resolve()),
            "-OutputFile",
            str(output_file.resolve()),
        ]
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=240)
        if proc.returncode != 0 or not output_file.exists():
            detail = (proc.stderr or proc.stdout or "Microsoft Word 自动化失败").strip()
            raise RuntimeError(
                "无法通过 Microsoft Word 完成 PDF/Word 转换。"
                "请确认已安装桌面版 Microsoft Word，并能正常打开该文件。\n" + detail[:1200]
            )


def _roundtrip_word_transformer(texts: Sequence[str], regex: bool, auto: bool):
    """Remove converted-PDF watermark objects and short text paragraphs."""
    base = core.word_transformer(texts, regex, auto, True)

    def transform(name: str, data: bytes):
        data2, changes = base(name, data)
        if name != "word/document.xml":
            return data2, changes
        try:
            root = core.parse_xml(data2)
        except Exception:
            return data2, changes

        # PDF reflow can turn a watermark into an ordinary short paragraph or
        # text-box paragraph. Remove only explicit/known watermark text.
        for para in list(root.xpath(".//w:p", namespaces=core.NS)):
            text = "".join(x.text or "" for x in para.xpath(".//w:t", namespaces=core.NS)).strip()
            if not text or len(text) > 140:
                continue
            explicit = core.text_matches(text, texts, regex)
            known = auto and core.looks_like_watermark_text(text)
            if explicit or known:
                parent = para.getparent()
                if parent is not None:
                    parent.remove(para)
                    changes.append(core.Change(name, "PDF转Word文字水印", text[:160]))

        if changes:
            return core.xml_bytes(root), changes
        return data2, changes

    return transform


def process_pdf_roundtrip(
    src: Path,
    output: Path | None,
    texts: Sequence[str],
    regex: bool,
    auto: bool,
    aggressive: bool,
    fallback_direct: bool = True,
) -> core.Result:
    """Process a PDF through Word and return the same Result model as the core."""
    src = src.resolve()
    out = (output.resolve() if output else core.default_output(src))
    warnings: list[str] = []

    discovered = discover_pdf_watermark_texts(src) if auto else []
    targets = list(texts)
    existing = {_norm(x) for x in targets}
    for item in discovered:
        if _norm(item) not in existing:
            targets.append(item)
            existing.add(_norm(item))
    if discovered:
        warnings.append("PDF 预识别到疑似水印文字：" + " | ".join(discovered[:8]))

    try:
        with tempfile.TemporaryDirectory(prefix="wm_pdf_roundtrip_") as td:
            work = Path(td)
            raw_docx = work / "converted.docx"
            clean_docx = work / "clean.docx"
            temp_pdf = work / "clean.pdf"

            _run_word_com(src, raw_docx, "pdf_to_docx")
            changes = core.rewrite_zip(
                raw_docx,
                clean_docx,
                _roundtrip_word_transformer(targets, regex, auto),
            )
            _run_word_com(clean_docx, temp_pdf, "docx_to_pdf")
            out.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(temp_pdf, out)

        warnings.append("PDF 已按“PDF→Word→去水印→PDF”流程处理；复杂排版可能发生轻微重排，请核对输出。")
        if not changes:
            warnings.append("Word 层未识别到明确可删除的水印对象；转换过程可能仍改变/忽略部分 PDF 图层，请人工确认水印是否已去除。")
        return core.Result(str(src), str(out), None, changes, warnings)
    except Exception as exc:
        if not fallback_direct:
            return core.Result(str(src), None, None, [], warnings, error=str(exc))
        warnings.append("Word 往返模式失败，已回退为 PDF 直接处理：" + str(exc))
        direct = core.process_one(src, out, texts, regex, auto, aggressive, False, None, False)
        direct.warnings = warnings + direct.warnings
        return direct
