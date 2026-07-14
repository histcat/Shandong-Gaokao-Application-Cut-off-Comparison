# Application-Entrance Score Comparison

> Compare your college application application form against the admission cutoff scores — instantly see which choice admits you.

A Python desktop tool that parses a Gaokao application form (PDF) and admission cutoff data (XLS), then compares your rank against each choice's minimum rank to determine admission results.

The format of the cutoff data (XLS) should follow the [山东省2025年普通类常规批第1次志愿投档情况表](https://www.sdzk.cn/NewsInfo.aspx?NewsID=6996).

The format of the college application form should follow the PDF file downloaded from the Application System.

## Effortless Access

Simply download the binary from the Releases page.

*Note: Currently only supports amd64 systems. For other architectures, please build from source.*

## Motivation

The Shandong Provincial Department of Education always announces the application status table a few days before the official admission results are released. You can refer to that table to check your result in advance.

To make the process more convenient, I built this program with the help of Claude Code.


## Features

- **Drag & Drop** — drag PDF and XLS files directly into the window
- **Material Design 3** — clean, modern UI with VS Code-style dark log panel
- **Dual Mode** — GUI for daily use, CLI for scripting/automation
- **Auto-save** — comparison log automatically saved as `comparison_log_<timestamp>.txt`
- **High-DPI Aware** — sharp rendering on 4K and high-DPI displays
- **CJK Encoding Safe** — built-in `sanitize()` handles CJK compatibility characters to prevent encoding errors

## Installation

### Prerequisites

- **Python 3.9+** (3.11 recommended)
- **OS:** Windows 10/11, macOS, or Linux

### Install Dependencies

```bash
pip install pandas PyPDF2 openpyxl xlrd tkinterdnd2
```

| Package | Purpose |
|---------|---------|
| `pandas` + `openpyxl` + `xlrd` | Read cutoff data (.xls/.xlsx) |
| `PyPDF2` | Parse application form PDF |
| `tkinterdnd2` | Drag-and-drop file support |
| `tkinter` | GUI toolkit (included with Python) |

> **Linux users:** `tkinter` may need separate installation:
> ```bash
> # Ubuntu/Debian
> sudo apt-get install python3-tk
> # Fedora
> sudo dnf install python3-tkinter
> # Arch
> sudo pacman -S tk
> ```

## Quick Start

1. **Prepare files** — place your `志愿表.pdf` and `投档线.xls` in the project directory.

2. **Launch the GUI:**
   ```bash
   python gui.py
   ```

3. **Drag files** into the drop zones (or click to browse).

4. **Enter your rank** (the smaller the number, the better).

5. **Click "开始比对"** (Start Comparison) — results appear instantly with a popup.

### CLI Version

```bash
python compare.py
```

Make sure `志愿表.pdf` and `投档线.xls` are in the same directory, or edit the `PDF_PATH` / `XLS_PATH` constants at the top of the script.

## Usage

```
┌──────────────────────────────────────────────┐
│  志愿-投档线对照分析                           │
│  ──────────────────────────────────────────── │
│                                              │
│  ┌──────────────┐   ┌──────────────┐         │
│  │  Application PDF │   │  Cutoff XLS  │        │
│  │  📂            │   │  📊           │        │
│  │  Drop or click │   │  Drop or click│        │
│  └──────────────┘   └──────────────┘         │
│                                              │
│  Enter your rank:                            │
│  ┌──────────────────────────────────┐        │
│  │  Rank (integer, lower is better) │        │
│  │  ______________                  │        │
│  └──────────────────────────────────┘        │
│                                              │
│        [ Start Comparison ]                   │
│                                              │
│  ┌─ Comparison Log ─────────────────┐        │
│  │ Choice 1: [A246] Fudan Univ.     │        │
│  │   [Not admitted] 150 > 248       │        │
│  │ Choice 2: ...                    │        │
│  │ Choice 7: [A246] Fudan Univ.     │        │
│  │   [Admitted] 150 <= 583         │        │
│  └──────────────────────────────────┘        │
└──────────────────────────────────────────────┘
```

## Project Structure

```
志愿表-投档线对照/
├── gui.py                    ← GUI application (Material Design 3)
├── compare.py                ← CLI version
├── cloud.ico                 ← Application icon
├── cloud.png                 ← Application icon (PNG fallback)
├── 志愿-投档线对照.spec       ← PyInstaller spec file
├── BUILD.md                  ← Build & packaging guide (Chinese)
├── README.md                 ← This file
├── README.zh.md              ← Chinese README
├── 志愿表.pdf                 ← (user-provided) Application form
├── 投档线.xls                 ← (user-provided) Cutoff data
└── comparison_log_*.txt      ← (auto-generated) Comparison logs
```

## Comparison Logic

```
For each row in the application form:
  Step 1: Extract school_code (e.g. A246) and major_code (e.g. 33)
  Step 2: Look up (school_code, major_code) in cutoff data
  Step 3: If found:
            Your rank ≤ minimum rank → ADMITTED! Stop.
            Your rank > minimum rank → Not admitted, continue.
  Step 4: If not found:
            Warn "code combination not found" (codes may have changed across years), continue.
```

- **Rank** (`位次`): Your province-wide ranking number. Lower = better.
- **Minimum Rank** (`最低位次`): The rank of the last admitted student in that major last year.
- **Rule:** Your rank ≤ minimum rank → likely admitted.

## Packaging as Standalone EXE

Build a single `.exe` file that runs without Python:

```bash
pip install pyinstaller pillow
```

```bash
pyinstaller \
  --name "志愿-投档线对照" \
  --windowed \
  --onefile \
  --icon cloud.ico \
  --add-data "cloud.ico;." \
  --add-data "cloud.png;." \
  --hidden-import tkinterdnd2 \
  --hidden-import tkinterdnd2.tkdnd \
  --hidden-import pandas \
  --hidden-import PyPDF2 \
  --hidden-import openpyxl \
  --hidden-import xlrd \
  --collect-all tkinterdnd2 \
  gui.py
```

> **PowerShell users:** Replace `\` with `` ` `` for line continuation.

The output will be at `dist/志愿-投档线对照.exe`.

### Common Build Issues

- **Large EXE (200MB+):** pandas bundles many unused modules. Slim down with `--exclude-module matplotlib scipy bokeh sqlalchemy` and use `--onedir` for faster startup.
- **Drag-and-drop broken after packaging:** Ensure the `tkdnd` library directory is included with `--add-data`.
- **Antivirus false positive:** Single-file EXEs commonly trigger AV. Use `--onedir` to reduce false positives.

## Troubleshooting

| Problem | Solution |
|---------|----------|
| `No module named 'tkinter'` | Install `python3-tk` (Linux) |
| Drag-and-drop not working | Click the drop zone to browse files instead |
| `UnicodeEncodeError` | Run `set PYTHONIOENCODING=utf-8` before launching |
| Chinese characters garbled | Ensure CJK fonts (e.g. Microsoft YaHei) are installed |

## Data Format

### Application Form (PDF)

The tool extracts entries containing:
- application number (志愿号)
- School code + name (e.g. `A246复旦大学`)
- Major code + name (e.g. `33计算机科学与技术`)
- School type (公办/民办/独立学院 etc.)

### Cutoff Data (XLS)

The Excel file should have columns:
- **专业代码及名称** — Major code + name (e.g. `33计算机科学与技术`)
- **院校代码及名称** — School code + name (e.g. `A246复旦大学`)
- **投档计划数** — Planned enrollment
- **最低位次** — Minimum admission rank

## License

[MIT](LICENSE)

---

Built with Python, tkinter, and Material Design 3.
