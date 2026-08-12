#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Office / PDF watermark remover

Supported directly:
  - Word OOXML: .docx, .docm
  - Excel OOXML: .xlsx, .xlsm
  - PowerPoint OOXML: .pptx, .pptm
  - OpenDocument: .odt, .ods, .odp (best effort)
  - PDF: .pdf (annotation watermarks; embedded text redaction is opt-in)

Legacy / native office formats via LibreOffice conversion (best effort):
  - .doc, .xls, .ppt, .wps, .et, .dps

Design goals:
  - Safe by default: remove only high-confidence watermark objects.
  - Preserve document formatting: manipulate package XML directly instead of re-saving
    through python-docx/openpyxl.
  - For PDF, do NOT blindly redact page text unless --pdf-redact-text is explicitly used,
    because redaction can also remove normal text overlapped by a watermark.

Examples:
  python office_watermark_remover.py report.docx
  python office_watermark_remover.py report.docx --text "顾磊" --text "2026-08-12"
  python office_watermark_remover.py folder --recursive
  python office_watermark_remover.py file.pdf --text "内部资料" --pdf-redact-text
  python office_watermark_remover.py old.wps --soffice "C:\\Program Files\\LibreOffice\\program\\soffice.exe"

Run without arguments to open a simple file-picker GUI.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Iterable, Optional, Sequence

try:
    from lxml import etree
except ImportError:
    etree = None


SUPPORTED_DIRECT = {".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm", ".odt", ".ods", ".odp", ".pdf"}
LEGACY_CONVERT = {
    ".doc": ".docx",
    ".wps": ".docx",
    ".xls": ".xlsx",
    ".et": ".xlsx",
    ".ppt": ".pptx",
    ".dps": ".pptx",
}
SUPPORTED_ALL = SUPPORTED_DIRECT | set(LEGACY_CONVERT)

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "svg": "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
}

WATERMARK_WORDS = (
    "watermark", "powerpluswatermarkobject", "水印", "机密", "秘密", "绝密",
    "内部资料", "内部使用", "仅供内部", "严禁复制", "样稿", "草稿", "draft",
    "confidential", "do not copy", "sample",
)


@dataclass
class Change:
    part: str
    kind: str
    detail: str


@dataclass
class Result:
    input: str
    output: Optional[str]
    converted_from: Optional[str]
    changes: list[Change]
    warnings: list[str]
    error: Optional[str] = None

    @property
    def count(self) -> int:
        return len(self.changes)


def require_lxml() -> None:
    if etree is None:
        raise RuntimeError("缺少依赖 lxml。请运行: pip install lxml pymupdf")


def norm_text(s: str) -> str:
    return re.sub(r"\s+", "", s or "").lower()


def text_matches(s: str, texts: Sequence[str], regex: bool = False) -> bool:
    if not texts:
        return False
    if regex:
        return any(re.search(p, s or "", re.I) for p in texts)
    ns = norm_text(s)
    return any(norm_text(t) in ns for t in texts if t)


def looks_like_watermark_text(s: str) -> bool:
    ns = norm_text(s)
    return any(norm_text(w) in ns for w in WATERMARK_WORDS)


def all_text(el) -> str:
    vals = []
    # visible text nodes
    for xp in (".//w:t", ".//a:t", ".//text:p"):
        try:
            vals.extend([x.text or "" for x in el.xpath(xp, namespaces=NS)])
        except Exception:
            pass
    # VML text path and common names / descriptions
    attrs = (
        "string", "id", "name", "descr", "title", "type", "style",
        f"{{{NS['draw']}}}name", f"{{{NS['svg']}}}desc",
    )
    for node in el.iter():
        for a in attrs:
            if a in node.attrib:
                vals.append(node.attrib.get(a, ""))
        # namespaced id/title variants
        for k, v in node.attrib.items():
            if k.endswith("}name") or k.endswith("}title") or k.endswith("}descr"):
                vals.append(v)
    return " ".join(vals)


def attr_blob(el) -> str:
    vals = []
    for node in el.iter():
        for k, v in node.attrib.items():
            key = k.rsplit("}", 1)[-1] if "}" in k else k
            vals.append(f"{key}={v}")
    return " ".join(vals).lower()


def remove_node(node) -> bool:
    parent = node.getparent()
    if parent is None:
        return False
    parent.remove(node)
    return True


def remove_best_word_container(node) -> bool:
    """Remove the smallest Word container that cleanly owns a watermark object."""
    cur = node
    q_pict = f"{{{NS['w']}}}pict"
    q_drawing = f"{{{NS['w']}}}drawing"
    while cur is not None:
        if cur.tag in (q_pict, q_drawing):
            return remove_node(cur)
        cur = cur.getparent()
    return remove_node(node)


def xml_bytes(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=None)


def parse_xml(data: bytes):
    parser = etree.XMLParser(remove_blank_text=False, recover=True, huge_tree=True)
    return etree.fromstring(data, parser=parser)


def rewrite_zip(src: Path, dst: Path, transform) -> list[Change]:
    require_lxml()
    changes: list[Change] = []
    dst.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dst, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            new_data, part_changes = transform(info.filename, data)
            changes.extend(part_changes)
            # preserve ZipInfo metadata as much as possible
            zout.writestr(info, new_data)
    return changes


def word_transformer(texts: Sequence[str], regex: bool, auto: bool, aggressive: bool):
    def transform(name: str, data: bytes):
        changes: list[Change] = []
        is_header = bool(re.fullmatch(r"word/header\d+\.xml", name))
        is_footer = bool(re.fullmatch(r"word/footer\d+\.xml", name))
        is_body = name == "word/document.xml"
        if not (is_header or is_footer or (aggressive and is_body)):
            return data, changes
        try:
            root = parse_xml(data)
        except Exception:
            return data, changes

        # VML shapes (classic Word/WPS watermark implementation)
        for shape in list(root.xpath(".//v:shape", namespaces=NS)):
            s = all_text(shape)
            blob = attr_blob(shape)
            explicit = text_matches(s, texts, regex)
            named = looks_like_watermark_text(s)
            positioned = (
                "position:absolute" in blob and
                ("mso-position-horizontal:center" in blob or "mso-position-vertical:center" in blob)
            )
            rotated = "rotation:" in blob or "rotation=" in blob
            textpath = bool(shape.xpath(".//v:textpath", namespaces=NS))
            high_conf = named or (is_header and textpath and positioned and rotated)
            if explicit or (auto and high_conf) or (aggressive and is_header and textpath and rotated):
                detail = s.strip()[:160] or "VML shape"
                if remove_best_word_container(shape):
                    changes.append(Change(name, "Word VML水印", detail))

        # DrawingML / WordprocessingShape watermarks
        for drawing in list(root.xpath(".//w:drawing", namespaces=NS)):
            s = all_text(drawing)
            blob = attr_blob(drawing)
            explicit = text_matches(s, texts, regex)
            named = looks_like_watermark_text(s)
            rotated = "rot" in blob or "rotation" in blob
            behind = "behinddoc" in blob and ("1" in blob or "true" in blob)
            centered = "positionh" in blob or "positionv" in blob
            high_conf = named or (is_header and rotated and (behind or centered))
            if explicit or (auto and high_conf) or (aggressive and is_header and rotated):
                detail = s.strip()[:160] or "DrawingML shape"
                if remove_node(drawing):
                    changes.append(Change(name, "Word DrawingML水印", detail))

        if changes:
            return xml_bytes(root), changes
        return data, changes
    return transform


def excel_transformer(texts: Sequence[str], regex: bool, auto: bool, aggressive: bool):
    def transform(name: str, data: bytes):
        changes: list[Change] = []
        if not name.endswith((".xml", ".vml")):
            return data, changes
        relevant = (
            re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name)
            or re.fullmatch(r"xl/drawings/drawing\d+\.xml", name)
            or re.fullmatch(r"xl/drawings/vmlDrawing\d+\.vml", name)
        )
        if not relevant:
            return data, changes
        try:
            root = parse_xml(data)
        except Exception:
            return data, changes

        # Header/footer graphics (&G is Excel's picture marker)
        if name.startswith("xl/worksheets/"):
            for hf in root.xpath(".//x:headerFooter", namespaces=NS):
                touched = False
                for child in list(hf):
                    txt = child.text or ""
                    new = txt
                    if text_matches(txt, texts, regex):
                        # remove requested text; if the entire field is a match, clear it
                        if regex:
                            for p in texts:
                                new = re.sub(p, "", new, flags=re.I)
                        else:
                            for t in texts:
                                new = re.sub(re.escape(t), "", new, flags=re.I)
                    if auto and "&G" in new:
                        new = new.replace("&G", "")
                    if new != txt:
                        child.text = new
                        touched = True
                        changes.append(Change(name, "Excel页眉/页脚水印", txt[:160]))
                # aggressive: remove sheet background picture node
                if aggressive:
                    for pic in list(root.xpath(".//x:picture", namespaces=NS)):
                        if remove_node(pic):
                            changes.append(Change(name, "Excel工作表背景图", "sheet background picture"))

        # Spreadsheet DrawingML shapes
        if name.startswith("xl/drawings/drawing"):
            for sp in list(root.xpath(".//xdr:sp", namespaces=NS)):
                s = all_text(sp)
                blob = attr_blob(sp)
                explicit = text_matches(s, texts, regex)
                named = looks_like_watermark_text(s)
                rotated = "rot" in blob
                alpha_vals = [int(v) for v in re.findall(r"alpha[^0-9]*(\d{3,6})", blob) if v.isdigit()]
                translucent = any(v < 80000 for v in alpha_vals) or "alpha" in blob
                high_conf = named or (rotated and translucent and bool(s.strip()))
                if explicit or (auto and high_conf) or (aggressive and rotated and bool(s.strip())):
                    if remove_node(sp):
                        changes.append(Change(name, "Excel绘图水印", s.strip()[:160] or "shape"))
            for pic in list(root.xpath(".//xdr:pic", namespaces=NS)):
                s = all_text(pic)
                if text_matches(s, texts, regex) or looks_like_watermark_text(s):
                    if remove_node(pic):
                        changes.append(Change(name, "Excel图片水印", s.strip()[:160] or "picture"))

        # Legacy VML shapes in Excel/WPS
        if name.endswith(".vml"):
            for shape in list(root.xpath(".//v:shape", namespaces=NS)):
                s = all_text(shape)
                blob = attr_blob(shape)
                explicit = text_matches(s, texts, regex)
                named = looks_like_watermark_text(s)
                rotated = "rotation" in blob
                translucent = "opacity" in blob
                if explicit or (auto and (named or (rotated and translucent))) or (aggressive and rotated):
                    if remove_node(shape):
                        changes.append(Change(name, "Excel/WPS VML水印", s.strip()[:160] or "VML shape"))

        if changes:
            return xml_bytes(root), changes
        return data, changes
    return transform


def ppt_transformer(texts: Sequence[str], regex: bool, auto: bool, aggressive: bool):
    def transform(name: str, data: bytes):
        changes: list[Change] = []
        if not name.endswith(".xml"):
            return data, changes
        relevant = bool(re.fullmatch(r"ppt/(slides/slide|slideMasters/slideMaster|slideLayouts/slideLayout)\d+\.xml", name))
        if not relevant:
            return data, changes
        try:
            root = parse_xml(data)
        except Exception:
            return data, changes

        for tag, kind in (("p:sp", "PPT文字/形状水印"), ("p:pic", "PPT图片水印")):
            for el in list(root.xpath(f".//{tag}", namespaces=NS)):
                s = all_text(el)
                blob = attr_blob(el)
                explicit = text_matches(s, texts, regex)
                named = looks_like_watermark_text(s)
                rotated = bool(re.search(r"\brot=\"?[1-9]", blob)) or " rot=" in blob
                translucent = "alpha" in blob
                is_text_shape = tag == "p:sp" and bool(s.strip())
                high_conf = named or (is_text_shape and rotated and translucent)
                if explicit or (auto and high_conf) or (aggressive and is_text_shape and rotated):
                    if remove_node(el):
                        changes.append(Change(name, kind, s.strip()[:160] or "shape/picture"))

        if changes:
            return xml_bytes(root), changes
        return data, changes
    return transform


def odf_transformer(texts: Sequence[str], regex: bool, auto: bool, aggressive: bool):
    def transform(name: str, data: bytes):
        changes: list[Change] = []
        if name not in {"content.xml", "styles.xml"}:
            return data, changes
        try:
            root = parse_xml(data)
        except Exception:
            return data, changes

        for xp, kind in ((".//draw:frame", "ODF框架水印"), (".//draw:custom-shape", "ODF形状水印")):
            for el in list(root.xpath(xp, namespaces=NS)):
                s = all_text(el)
                blob = attr_blob(el)
                explicit = text_matches(s, texts, regex)
                named = looks_like_watermark_text(s)
                rotated = "rotate" in blob or "rotation" in blob or "transform" in blob
                translucent = "opacity" in blob or "transparency" in blob
                high_conf = named or (rotated and translucent and bool(s.strip()))
                if explicit or (auto and high_conf) or (aggressive and rotated and bool(s.strip())):
                    if remove_node(el):
                        changes.append(Change(name, kind, s.strip()[:160] or "shape"))
        if changes:
            return xml_bytes(root), changes
        return data, changes
    return transform


def remove_pdf(src: Path, dst: Path, texts: Sequence[str], regex: bool, auto: bool,
               aggressive: bool, pdf_redact_text: bool) -> tuple[list[Change], list[str]]:
    try:
        import pymupdf
    except ImportError as e:
        raise RuntimeError("处理 PDF 需要 PyMuPDF。请运行: pip install pymupdf") from e

    changes: list[Change] = []
    warnings: list[str] = []
    doc = pymupdf.open(src)

    # 1) Annotation watermarks / stamps: safe removal because they are separate objects.
    for page in doc:
        annots = list(page.annots() or [])
        for annot in annots:
            info = annot.info or {}
            info_text = " ".join(str(v) for v in info.values() if v)
            try:
                raw = doc.xref_object(annot.xref, compressed=False)
            except Exception:
                raw = ""
            type_name = (annot.type[1] if isinstance(annot.type, tuple) and len(annot.type) > 1 else str(annot.type)) or ""
            blob = f"{info_text} {raw} {type_name}"
            explicit = text_matches(blob, texts, regex)
            named = looks_like_watermark_text(blob) or "/Subtype /Watermark" in raw
            stampish = type_name.lower() in {"stamp", "freeText".lower(), "watermark"}
            if explicit or (auto and named) or (aggressive and stampish):
                page.delete_annot(annot)
                changes.append(Change(f"page {page.number+1}", "PDF批注/印章水印", (info_text or type_name)[:160]))

    # 2) Embedded page text. Opt-in only because redaction removes any text overlapping the same rectangle.
    if pdf_redact_text:
        if not texts:
            warnings.append("已指定 --pdf-redact-text，但没有 --text；跳过 PDF 正文文字删除。")
        elif regex:
            warnings.append("PDF 的 --pdf-redact-text 目前只支持普通文本搜索，不支持正则；将把 --text 当作字面文本。")
        for page in doc:
            rects = []
            for t in texts:
                if not t:
                    continue
                try:
                    rects.extend(page.search_for(t))
                except Exception:
                    pass
            for rect in rects:
                # No fill; keep images and vector graphics. Any text glyph intersecting this rect can still be removed.
                page.add_redact_annot(rect, fill=None, cross_out=False)
            if rects:
                try:
                    page.apply_redactions(
                        images=pymupdf.PDF_REDACT_IMAGE_NONE,
                        graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
                        text=pymupdf.PDF_REDACT_TEXT_REMOVE,
                    )
                except TypeError:
                    # Compatibility fallback for older versions.
                    page.apply_redactions(images=0, graphics=0)
                for r in rects:
                    changes.append(Change(f"page {page.number+1}", "PDF正文文字删除", f"rect={tuple(round(x,2) for x in r)}"))
        if any(c.kind == "PDF正文文字删除" for c in changes):
            warnings.append("PDF 正文水印采用文字红删方式：若水印与正文文字重叠，重叠正文也可能被删除，请务必核对输出文件。")

    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(dst, garbage=4, deflate=True, clean=True)
    doc.close()
    return changes, warnings


def find_soffice(user_path: Optional[str] = None) -> Optional[str]:
    candidates = []
    if user_path:
        candidates.append(user_path)
    for cmd in ("soffice", "libreoffice"):
        p = shutil.which(cmd)
        if p:
            candidates.append(p)
    if os.name == "nt":
        candidates.extend([
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ])
    for p in candidates:
        if p and Path(p).exists():
            return str(Path(p))
    return None


def convert_legacy(src: Path, target_ext: str, soffice: Optional[str]) -> Path:
    exe = find_soffice(soffice)
    if not exe:
        raise RuntimeError(
            f"{src.suffix} 需要先转换为现代 Office 格式。未找到 LibreOffice。"
            "请安装 LibreOffice，或使用 --soffice 指定 soffice.exe；"
            "如果是 WPS 原生 .wps/.et/.dps 且 LibreOffice 无法识别，请在 WPS 中另存为 docx/xlsx/pptx 后再运行。"
        )
    outdir = Path(tempfile.mkdtemp(prefix="wm_convert_"))
    target = target_ext.lstrip(".")
    cmd = [exe, "--headless", "--convert-to", target, "--outdir", str(outdir), str(src)]
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=180)
    expected = outdir / f"{src.stem}.{target}"
    if proc.returncode != 0 or not expected.exists():
        raise RuntimeError(
            "LibreOffice 转换失败。\n"
            f"命令输出: {proc.stdout.strip()}\n错误输出: {proc.stderr.strip()}\n"
            "WPS 原生格式的兼容性依赖本机转换器，建议先在 WPS 中另存为 docx/xlsx/pptx。"
        )
    return expected


def default_output(src: Path, actual_ext: Optional[str] = None) -> Path:
    ext = actual_ext or src.suffix
    return src.with_name(src.stem + "_无水印" + ext)


def process_one(src: Path, output: Optional[Path], texts: Sequence[str], regex: bool,
                auto: bool, aggressive: bool, pdf_redact_text: bool,
                soffice: Optional[str], dry_run: bool = False) -> Result:
    src = src.resolve()
    warnings: list[str] = []
    converted_from = None
    work_src = src
    ext = src.suffix.lower()

    try:
        if ext not in SUPPORTED_ALL:
            raise RuntimeError(f"暂不支持该格式: {ext}")

        if ext in LEGACY_CONVERT:
            converted_from = str(src)
            target_ext = LEGACY_CONVERT[ext]
            work_src = convert_legacy(src, target_ext, soffice)
            ext = target_ext
            warnings.append(f"原格式 {src.suffix} 已先转换为 {target_ext} 再处理；输出将使用现代格式 {target_ext}。")

        if output is None:
            out = default_output(src, ext if converted_from else None)
        else:
            out = output.resolve()
            if converted_from and out.suffix.lower() != ext:
                out = out.with_suffix(ext)
                warnings.append(f"输出扩展名已调整为 {ext}。")

        if dry_run:
            tmp_out = Path(tempfile.mkstemp(suffix=ext, prefix="wm_dry_")[1])
        else:
            tmp_out = out

        if ext in {".docx", ".docm"}:
            changes = rewrite_zip(work_src, tmp_out, word_transformer(texts, regex, auto, aggressive))
        elif ext in {".xlsx", ".xlsm"}:
            changes = rewrite_zip(work_src, tmp_out, excel_transformer(texts, regex, auto, aggressive))
        elif ext in {".pptx", ".pptm"}:
            changes = rewrite_zip(work_src, tmp_out, ppt_transformer(texts, regex, auto, aggressive))
        elif ext in {".odt", ".ods", ".odp"}:
            changes = rewrite_zip(work_src, tmp_out, odf_transformer(texts, regex, auto, aggressive))
        elif ext == ".pdf":
            changes, pdf_warnings = remove_pdf(work_src, tmp_out, texts, regex, auto, aggressive, pdf_redact_text)
            warnings.extend(pdf_warnings)
        else:
            raise RuntimeError(f"转换后的格式仍不支持: {ext}")

        if dry_run:
            try:
                tmp_out.unlink(missing_ok=True)
            except Exception:
                pass
            final_output = None
        else:
            final_output = str(out)

        if not changes:
            warnings.append("未发现满足当前规则的高置信度水印对象。可尝试 --text 指定水印文字，或谨慎使用 --aggressive。")

        return Result(str(src), final_output, converted_from, changes, warnings)
    except Exception as e:
        return Result(str(src), None, converted_from, [], warnings, error=str(e))


def iter_inputs(path: Path, recursive: bool) -> Iterable[Path]:
    if path.is_file():
        yield path
        return
    if path.is_dir():
        it = path.rglob("*") if recursive else path.glob("*")
        for p in it:
            if p.is_file() and p.suffix.lower() in SUPPORTED_ALL and "_无水印" not in p.stem:
                yield p


def print_result(r: Result) -> None:
    print(f"\n[文件] {r.input}")
    if r.error:
        print(f"  [失败] {r.error}")
        return
    print(f"  [完成] 删除/修改 {r.count} 个水印相关对象")
    for c in r.changes:
        print(f"    - {c.kind}: {c.part} | {c.detail}")
    for w in r.warnings:
        print(f"  [提示] {w}")
    if r.output:
        print(f"  [输出] {r.output}")


def run_gui() -> int:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, simpledialog
    except Exception:
        print("当前环境无法启动图形界面，请从命令行运行并传入文件路径。")
        return 2

    root = tk.Tk()
    root.withdraw()
    files = filedialog.askopenfilenames(
        title="选择需要去水印的文件（可多选）",
        filetypes=[
            ("办公文件", "*.docx *.docm *.xlsx *.xlsm *.pptx *.pptm *.pdf *.odt *.ods *.odp *.doc *.xls *.ppt *.wps *.et *.dps"),
            ("所有文件", "*.*"),
        ],
    )
    if not files:
        return 0
    text = simpledialog.askstring(
        "可选：指定水印文字",
        "如知道水印文字，可输入一部分（例如“顾磊”或“内部资料”）。\n不知道可留空，程序使用安全自动识别。",
        parent=root,
    ) or ""
    texts = [text] if text else []
    results = [process_one(Path(f), None, texts, False, True, False, False, None) for f in files]
    ok = sum(1 for r in results if not r.error)
    changed = sum(r.count for r in results if not r.error)
    outputs = "\n".join(r.output for r in results if r.output)[:2500]
    errors = "\n".join(f"{r.input}: {r.error}" for r in results if r.error)[:1500]
    msg = f"处理完成：{ok}/{len(results)} 个文件成功，共删除/修改 {changed} 个水印对象。\n\n输出：\n{outputs}"
    if errors:
        msg += f"\n\n失败：\n{errors}"
    messagebox.showinfo("去水印完成", msg, parent=root)
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Word / Excel / PowerPoint / PDF / WPS兼容格式 批量去水印工具",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    p.add_argument("inputs", nargs="*", help="文件或目录路径；不传参数时打开图形界面")
    p.add_argument("-o", "--output", help="单文件输出路径（多文件/目录模式忽略）")
    p.add_argument("--text", action="append", default=[], help="指定水印文字，可重复使用，例如 --text 顾磊 --text 2026-08-12")
    p.add_argument("--regex", action="store_true", help="把 --text 当作正则表达式（PDF正文搜索除外）")
    p.add_argument("--no-auto", action="store_true", help="关闭自动高置信度识别，只按 --text 删除")
    p.add_argument("--aggressive", action="store_true", help="激进模式：扩大删除范围，可能误删普通页眉/图形，使用前建议备份")
    p.add_argument("--pdf-redact-text", action="store_true", help="允许删除PDF正文内容流中的指定文字；可能伤及与水印重叠的正文")
    p.add_argument("--recursive", action="store_true", help="目录模式递归处理子目录")
    p.add_argument("--soffice", help="LibreOffice soffice.exe 路径，用于 .doc/.xls/.ppt/.wps/.et/.dps 转换")
    p.add_argument("--dry-run", action="store_true", help="只扫描并报告，不保留输出文件")
    p.add_argument("--json", dest="json_path", help="把处理报告写入 JSON 文件")
    return p


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if not args.inputs:
        return run_gui()

    input_paths = [Path(x) for x in args.inputs]
    files: list[Path] = []
    for p in input_paths:
        files.extend(iter_inputs(p, args.recursive))
    if not files:
        print("没有找到可处理的文件。")
        return 2

    multi = len(files) > 1 or any(p.is_dir() for p in input_paths)
    results: list[Result] = []
    for f in files:
        out = None if multi or not args.output else Path(args.output)
        r = process_one(
            f, out, args.text, args.regex, not args.no_auto, args.aggressive,
            args.pdf_redact_text, args.soffice, args.dry_run,
        )
        results.append(r)
        print_result(r)

    if args.json_path:
        report = [
            {**asdict(r), "count": r.count}
            for r in results
        ]
        Path(args.json_path).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"\n报告已写入: {args.json_path}")

    failed = sum(1 for r in results if r.error)
    print(f"\n汇总：成功 {len(results)-failed}，失败 {failed}，共修改 {sum(r.count for r in results)} 个对象。")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
