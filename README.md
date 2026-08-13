# Office Watermark Remover

一个面向日常办公文件的批量去水印工具，支持 Word、Excel、PowerPoint、PDF、OpenDocument，以及部分 WPS/旧版 Office 格式。

> 建议仅用于你有权修改的文件。正式材料处理后应人工复核版式和内容完整性。

## 功能特点

- **Windows 图形界面**：支持文件选择、文件夹批量导入、拖拽、进度显示、处理日志和输出目录选择。
- **Windows 单文件 EXE**：可通过 GitHub Actions 或本地 PyInstaller 构建，无需在使用电脑上单独安装 Python。
- **安全默认模式**：只删除高置信度水印对象，尽量避免误删正常页眉、Logo、艺术字和正文。
- **Word / WPS 特殊水印**：可直接处理藏在页眉 XML、VML、DrawingML 中的水印对象，适合 WPS 中“删除水印”无法去掉的姓名、日期、斜向水印。
- **Excel**：可处理页眉页脚图片标记、绘图层/VML 形状水印；激进模式下可处理工作表背景图。
- **PowerPoint**：可处理幻灯片、母版、版式中的文字、形状和图片水印。
- **PDF**：默认只删除独立的 Watermark/Stamp/批注类对象；正文内容流中的文字水印需要显式启用文字红删。
- **批量处理**：支持单文件、目录及递归子目录。
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

## 最简单的用法：Windows EXE

仓库已经提供 GitHub Actions 自动构建。进入仓库的 **Actions → Build Windows EXE**，打开最近一次成功的构建，在页面底部下载：

```text
OfficeWatermarkRemover-Windows
```

解压后运行：

```text
OfficeWatermarkRemover.exe
```

图形版支持：

- 直接拖入多个文件或文件夹
- 批量去水印
- 指定水印关键词，多个关键词使用 `|` 分隔
- 安全自动识别
- 激进模式二次确认
- PDF 指定文字删除
- 只扫描不输出
- 自定义输出目录
- 实时进度、单文件状态和详细处理日志

默认不会覆盖原文件，而是生成带 `_无水印` 后缀的新文件。

## 本地生成 Windows EXE

Windows 用户可以直接双击：

```text
生成Windows_EXE.bat
```

脚本会自动安装依赖和 PyInstaller，并生成：

```text
dist\OfficeWatermarkRemover.exe
```

等价的命令为：

```powershell
python -m pip install -r requirements.txt
python -m pip install pyinstaller
python -m PyInstaller --noconfirm --clean --onefile --windowed --name OfficeWatermarkRemover --collect-all tkinterdnd2 office_watermark_gui.py
```

## Python 安装

需要 Python 3.10 或更高版本。

```powershell
python -m pip install -r requirements.txt
```

Windows 用户也可以直接双击：

```text
安装依赖.bat
```

## Python 图形界面

运行完整图形界面：

```powershell
python office_watermark_gui.py
```

也可以运行原来的轻量文件选择界面：

```powershell
python office_watermark_remover.py
```

## 命令行使用方法

### 单文件自动识别

```powershell
python office_watermark_remover.py "材料.docx"
```

输出文件默认在原目录，并增加 `_无水印` 后缀。

### 指定水印文字

```powershell
python office_watermark_remover.py "材料.docx" --text "张三" --text "2026-08-12"
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

## Windows 拖拽脚本

安装依赖后，也可把文件直接拖到：

```text
去水印_拖拽文件到这里.bat
```

脚本会调用命令行核心处理拖入的文件。

## Word 水印为什么 WPS 有时删不掉？

有些文档里的“水印”并不是 WPS 菜单中的标准页面水印，而是保存在页眉 XML 中的 VML/DrawingML 形状对象，例如 `PowerPlusWaterMarkObject`、旋转文本路径、透明艺术字等。

本工具直接读取 OOXML 压缩包内部结构，识别并删除这些水印对象，因此可以处理部分 WPS 常规“删除水印”功能无法识别的情况，同时避免重新排版整个文档。

## 安全策略

### 默认模式

默认以减少误删为优先，只删除高置信度水印对象。对 Word、Excel、PowerPoint 主要通过修改 OOXML 包内 XML 完成，不通过 Office/WPS 重新排版整个文件。

### 激进模式

激进模式会扩大形状、背景和页眉对象的识别范围。GUI 中开启时会再次确认。建议重要材料先使用“只扫描不输出”查看识别情况。

### PDF

PDF 默认只处理独立水印/印章/批注对象。已经压平进正文内容流的水印文字，必须同时填写指定文字并启用“PDF 删除指定正文文字”。如果水印与正文重叠，可能连带影响重叠文字。

## GitHub Actions

`.github/workflows/build-windows.yml` 会在以下情况下构建 Windows EXE：

- 手工点击 `workflow_dispatch`
- `main` 分支上的核心代码、GUI、依赖或构建工作流发生更新

构建使用 Windows Runner + Python 3.12 + PyInstaller，生成的 ZIP 作为 Actions Artifact 保存 30 天。

## 注意事项

- 默认模式以减少误删为优先，不保证识别所有水印。
- Word/Excel/PPT 采用直接修改包内 XML 的方式，通常比使用 Office 自动化重新保存更能保持原版式。
- PDF 如果水印已经被压平到图片或复杂矢量内容中，无法保证自动无损去除。
- `.wps/.et/.dps/.doc/.xls/.ppt` 仍依赖本机 LibreOffice 转换；EXE 本身不会内置 LibreOffice。
- 重要合同、财务、审计、正式公文等文件处理后请务必与原件比对。

## 依赖

运行依赖：

- `lxml`
- `PyMuPDF`
- `tkinterdnd2`（GUI 拖拽）

可选：

- LibreOffice（旧 Office/WPS 格式转换）
- PyInstaller（本地构建 Windows EXE）
