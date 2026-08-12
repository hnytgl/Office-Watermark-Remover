# Office Watermark Remover

一个面向日常办公文件的批量去水印工具，支持 Word、Excel、PowerPoint、PDF、OpenDocument，以及部分 WPS/旧版 Office 格式。

> 建议仅用于你有权修改的文件。正式材料处理后应人工复核版式和内容完整性。

## 功能特点

- **安全默认模式**：只删除高置信度水印对象，尽量避免误删正常页眉、Logo、艺术字和正文。
- **Word / WPS 特殊水印**：可直接处理藏在页眉 XML、VML、DrawingML 中的水印对象，适合 WPS 中“删除水印”无法去掉的姓名、日期、斜向水印。
- **Excel**：可处理页眉页脚图片标记、绘图层/VML 形状水印；激进模式下可处理工作表背景图。
- **PowerPoint**：可处理幻灯片、母版、版式中的文字、形状和图片水印。
- **PDF**：默认只删除独立的 Watermark/Stamp/批注类对象；正文内容流中的文字水印需要显式启用文字红删。
- **批量处理**：支持单文件、目录及递归子目录。
- **图形界面**：不传参数时可直接弹出文件选择框。
- **WPS/旧 Office 格式**：`.wps/.et/.dps/.doc/.xls/.ppt` 可尝试通过 LibreOffice 转换后处理。

## 支持格式

### 直接支持

- Word：`.docx`、`.docm`
- Excel：`.xlsx`、`.xlsm`
- PowerPoint：`.pptx`、`.pptm`
- PDF：`.pdf`
- OpenDocument：`.odt`、`.ods`、`.odp`

### 尝试转换后处理

- Word/WPS：`.doc`、`.wps`
- Excel/WPS：`.xls`、`.et`
- PowerPoint/WPS：`.ppt`、`.dps`

> WPS 原生 `.wps/.et/.dps` 的兼容性取决于本机转换器。如果转换失败，建议先在 WPS 中“另存为” `docx/xlsx/pptx` 后再运行。

## 安装

需要 Python 3.10 或更高版本。

```powershell
python -m pip install -r requirements.txt
```

Windows 用户也可以直接双击：

```text
安装依赖.bat
```

## 使用方法

### 图形界面

直接运行：

```powershell
python office_watermark_remover.py
```

不传参数时会弹出文件选择框，可一次选择多个文件。

### 单文件自动识别

```powershell
python office_watermark_remover.py "材料.docx"
```

输出文件默认在原目录，并增加 `_无水印` 后缀。

### 指定水印文字

```powershell
python office_watermark_remover.py "材料.docx" --text "顾磊" --text "2026-08-12"
```

当自动识别不到某些特殊水印时，指定水印文字通常更稳。

### 批量处理目录

```powershell
python office_watermark_remover.py "D:\材料" --recursive
```

### 只扫描，不生成输出文件

```powershell
python office_watermark_remover.py "材料.docx" --dry-run
```

### PDF 正文文字水印

```powershell
python office_watermark_remover.py "材料.pdf" --text "内部资料" --pdf-redact-text
```

PDF 正文文字水印采用 redaction 方式处理。若水印文字与正常正文重叠，重叠区域中的正文也可能受到影响，因此处理后必须人工核对。

### 激进模式

```powershell
python office_watermark_remover.py "材料.docx" --aggressive
```

`--aggressive` 会扩大识别范围，可能误删普通艺术字、背景图或印章，建议先备份，或先使用 `--dry-run` 检查。

## Windows 拖拽使用

安装依赖后，可把文件直接拖到：

```text
去水印_拖拽文件到这里.bat
```

脚本会自动调用 `office_watermark_remover.py` 处理拖入的文件。

## Word 水印为什么 WPS 有时删不掉？

有些文档里的“水印”并不是 WPS 菜单中的标准页面水印，而是保存在页眉 XML 中的 VML/DrawingML 形状对象，例如 `PowerPlusWaterMarkObject`、旋转文本路径、透明艺术字等。

本工具直接读取 OOXML 压缩包内部结构，识别并删除这些水印对象，因此可以处理部分 WPS 常规“删除水印”功能无法识别的情况，同时避免重新排版整个文档。

## 注意事项

- 默认模式以减少误删为优先，不保证识别所有水印。
- Word/Excel/PPT 采用直接修改包内 XML 的方式，通常比使用 Office 自动化重新保存更能保持原版式。
- PDF 如果水印已经被压平到图片或复杂矢量内容中，无法保证自动无损去除。
- 重要合同、财务、审计、正式公文等文件处理后请务必与原件比对。

## 依赖

- `lxml`
- `PyMuPDF`
- 可选：LibreOffice（用于旧 Office/WPS 格式转换）
