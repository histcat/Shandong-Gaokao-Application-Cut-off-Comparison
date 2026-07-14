# 志愿-投档线对照分析 — 构建与使用指南

## 环境要求

| 项目 | 说明 |
|------|------|
| Python | **3.9+** （推荐 3.11） |
| 操作系统 | Windows 10/11, macOS, Linux |
| 依赖库 | 见下方安装步骤 |

## 第一步：安装 Python 依赖

打开终端（PowerShell / CMD / Bash），安装以下包：

```bash
pip install pandas PyPDF2 openpyxl xlrd tkinterdnd2
```

> **说明：**
> - `pandas` + `openpyxl` + `xlrd` — 读取投档线 .xls/.xlsx 文件
> - `PyPDF2` — 解析志愿表 PDF
> - `tkinterdnd2` — 支持从文件管理器拖拽文件到窗口
>
> `tkinter` 是 Python 自带的标准库，无需额外安装。

## 第二步：准备文件

确保以下两个文件与 `gui.py` 在同一目录：

```
志愿表-投档线对照/
├── gui.py          ← 图形界面程序
├── compare.py      ← 命令行版本（可选）
├── 志愿表.pdf       ← 你的志愿表
├── 投档线.xls       ← 去年的投档线数据
└── BUILD.md        ← 本文件
```

## 第三步：启动程序

```bash
cd 志愿表-投档线对照
python gui.py
```

## 使用流程

```
┌─────────────────────────────────────────────┐
│  志愿-投档线对照分析                          │
│  ─────────────────────────────────────────  │
│                                             │
│  ┌──────────────┐   ┌──────────────┐        │
│  │  志愿表 PDF   │   │  投档线 XLS   │        │
│  │  📂           │   │  📊           │        │
│  │  拖拽或点击   │   │  拖拽或点击   │        │
│  │  选择文件     │   │  选择文件     │        │
│  └──────────────┘   └──────────────┘        │
│                                             │
│  输入你的位次                                │
│  ┌──────────────────────────────────┐       │
│  │  位次（整数，越小越靠前）          │       │
│  │  ______________                   │       │
│  └──────────────────────────────────┘       │
│                                             │
│        [ 开始比对 ]                          │
│                                             │
│  ┌─ 比对日志 ──────────────────────┐        │
│  │ 志愿 1: [A246] 复旦大学           │        │
│  │   [未录取] 150 > 248             │        │
│  │ 志愿 2: ...                      │        │
│  │ 志愿 7: [A246] 复旦大学           │        │
│  │   [录取] 150 <= 583             │        │
│  └──────────────────────────────────┘       │
└─────────────────────────────────────────────┘
```

### 操作步骤

1. **拖入文件**（或点击选择）
   - 左侧放入 `志愿表.pdf`
   - 右侧放入 `投档线.xls`

2. **输入位次**
   - 填入你的全省排名位次（例如 `150`）
   - 位次越小 = 排名越靠前

3. **点击"开始比对"**
   - 程序自动解析 PDF 和 XLS
   - 逐条比对志愿代码与投档线数据
   - 实时在日志区显示比对过程

4. **查看结果**
   - **录取** → 弹窗显示录取院校、专业、位次对比
   - **未录取** → 弹窗提示所有志愿均未达线
   - 完整日志自动保存为 `comparison_log_日期时间.txt`

## 比对逻辑说明

```
对于志愿表中的每一行:
  1. 提取 院校代码（如 A246）和 专业代码（如 33）
  2. 在投档线中查找 (院校代码, 专业代码) 组合
  3. 如果找到:
     用户位次 <= 最低位次 → 录取！停止
     用户位次 >  最低位次 → 不录取，继续下一条
  4. 如果未找到:
     提示"代码组合未找到"（跨年代码变化），继续下一条
```

- **位次**：全省排名，数字越小越好
- **最低位次**：去年该专业录取的最后一名学生的位次
- **判断规则**：你的位次 ≤ 最低位次 → 可能被录取

## 两种版本

| 版本 | 文件 | 适用场景 |
|------|------|---------|
| **图形界面** | `gui.py` | 日常使用，拖拽操作，弹窗提示 |
| **命令行** | `compare.py` | 脚本化/自动化，`python compare.py` |

## 打包为独立 EXE（无需 Python 环境）

### 前置准备

安装 PyInstaller：
```bash
pip install pyinstaller pillow
```

### 打包命令

在项目目录下执行：

```bash
pyinstaller ^
  --name "志愿-投档线对照" ^
  --windowed ^
  --onefile ^
  --icon cloud.ico ^
  --add-data "cloud.ico;." ^
  --add-data "cloud.png;." ^
  --hidden-import tkinterdnd2 ^
  --hidden-import tkinterdnd2.tkdnd ^
  --hidden-import pandas ^
  --hidden-import PyPDF2 ^
  --hidden-import openpyxl ^
  --hidden-import xlrd ^
  --collect-all tkinterdnd2 ^
  gui.py
```

> **PowerShell 用户**：把 `^` 换成 `` ` ``。

### 参数解释

| 参数 | 作用 |
|------|------|
| `--name` | 输出 exe 文件名 |
| `--windowed` | 不弹命令行黑窗（纯 GUI 模式） |
| `--onefile` | 打包成**单个 exe**，方便分发 |
| `--icon cloud.ico` | exe 文件图标（文件管理器里看到的图标） |
| `--add-data "cloud.ico;."` | 把 ico 和 png 捆进 exe，运行时窗口左上角图标 |
| `--hidden-import` | 显式声明 PyInstaller 检测不到的隐式导入 |
| `--collect-all tkinterdnd2` | 收集拖拽库的 Tcl/Tk 扩展 dll |

### 输出位置

```
dist\
└── 志愿-投档线对照.exe    ← 双击即可运行，发给别人也能直接用
```

### 常见打包问题

**Q: exe 体积过大（200MB+）**

pandas 附带了很多不需要的模块。瘦身方法：
```bash
--onedir           # 用文件夹代替单文件（启动更快）
--exclude-module matplotlib
--exclude-module scipy
--exclude-module PIL
--exclude-module bokeh
--exclude-module sqlalchemy
```

**Q: 拖拽功能在打包后失效**

把 tkdnd 库目录也一起打包：
```bash
# 先找到 tkdnd 位置
python -c "import tkinterdnd2; import os; print(os.path.dirname(tkinterdnd2.__path__[0]))"

# 然后追加
--add-data "C:\Users\admin\AppData\Local\Programs\Python\Python311\tcl\tkdnd2.8;tkdnd"
```

**Q: 杀毒软件报毒**

单文件 exe 容易触发误报，改用 `--onedir` 可大幅降低误报率。

## 故障排除

### Q: 启动报错 `No module named 'tkinter'`

Linux 系统需单独安装 tkinter：

```bash
# Ubuntu/Debian
sudo apt-get install python3-tk

# Fedora
sudo dnf install python3-tkinter

# Arch
sudo pacman -S tk
```

### Q: 拖拽文件没反应

可能是 `tkinterdnd2` 未正确安装，可以点击文件区域手动浏览选择文件。

### Q: 报错 `UnicodeEncodeError`

程序已内置 `sanitize()` 函数处理 CJK 兼容字符，如果仍有编码问题，尝试在终端执行前设置：

```bash
set PYTHONIOENCODING=utf-8
python gui.py
```

### Q: 日志中没有显示中文

GUI 使用暗色日志面板（VS Code 风格），中文正常显示。如果出现乱码，检查系统是否安装了中文字体（Microsoft YaHei）。

## 文件清单

```
志愿表-投档线对照/
├── gui.py                       ← 图形界面主程序
├── compare.py                   ← 命令行版本
├── BUILD.md                     ← 本指南
├── 志愿表.pdf                    ← (你提供) 考生志愿表
├── 投档线.xls                    ← (你提供) 山东省投档线
└── comparison_log_*.txt         ← (自动生成) 比对日志
```
