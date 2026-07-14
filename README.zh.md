# 志愿-投档线对照分析

> 将你的高考志愿表和投档线数据比对，一键查看哪个志愿会被录取。

一个 Python 桌面工具，解析志愿表 PDF 和投档线 XLS 文件，逐条比对位次，判断录取结果。

投档线 XLS 文件格式参考[山东省2025年普通类常规批第1次志愿投档情况表](https://www.sdzk.cn/NewsInfo.aspx?NewsID=6996)

志愿表 PDF 格式参考从报考系统中导出的 PDF 文件。

## 功能特点

- **拖拽操作** — 直接将 PDF 和 XLS 文件拖入窗口即可加载
- **Material Design 3** — 简洁现代的界面，搭配 VS Code 风格暗色日志面板
- **双模式** — GUI 图形界面日常使用，CLI 命令行版本适合脚本自动化
- **日志自动保存** — 比对日志自动保存为 `comparison_log_时间戳.txt`
- **高分屏适配** — 支持 4K 和高 DPI 显示器，字体清晰不模糊
- **编码兼容** — 内置 `sanitize()` 函数处理 CJK 兼容字符，避免 GBK 编码错误

## 开发动机

山东省教育厅通常会在正式录取结果公布前几天发布申请状态表。你可以参考该表提前查询录取结果。

为了方便这一查询过程，我在 Claude Code 的辅助下开发了这个程序。


## 环境要求

- **Python 3.9+**（推荐 3.11）
- **操作系统：** Windows 10/11、macOS 或 Linux

### 安装依赖

```bash
pip install pandas PyPDF2 openpyxl xlrd tkinterdnd2
```

| 包名 | 用途 |
|------|------|
| `pandas` + `openpyxl` + `xlrd` | 读取投档线 .xls/.xlsx 文件 |
| `PyPDF2` | 解析志愿表 PDF |
| `tkinterdnd2` | 拖拽文件到窗口 |
| `tkinter` | GUI 工具包（Python 自带） |

> **Linux 用户** 需要单独安装 tkinter：
> ```bash
> # Ubuntu/Debian
> sudo apt-get install python3-tk
> # Fedora
> sudo dnf install python3-tkinter
> # Arch
> sudo pacman -S tk
> ```

## 快速开始

1. **准备文件** — 将你的 `志愿表.pdf` 和 `投档线.xls` 放到项目目录下。

2. **启动图形界面：**
   ```bash
   python gui.py
   ```

3. **拖入文件** 到对应的拖放区域（也可以点击区域浏览选择）。

4. **输入你的位次**（全省排名，数字越小越好）。

5. **点击"开始比对"** — 结果即时弹窗显示，日志区域同步输出详细过程。

### 命令行版本

```bash
python compare.py
```

确保 `志愿表.pdf` 和 `投档线.xls` 在同一目录下，或修改脚本开头的 `PDF_PATH` / `XLS_PATH` 常量。

## 使用流程

```
┌──────────────────────────────────────────────┐
│  志愿-投档线对照分析                           │
│  ──────────────────────────────────────────── │
│                                              │
│  ┌──────────────┐   ┌──────────────┐         │
│  │  志愿表 PDF   │   │  投档线 XLS   │        │
│  │  📂           │   │  📊           │        │
│  │  拖拽或点击   │   │  拖拽或点击   │        │
│  └──────────────┘   └──────────────┘         │
│                                              │
│  输入你的位次：                                │
│  ┌──────────────────────────────────┐        │
│  │  位次（整数，越小越靠前）          │        │
│  │  ______________                   │        │
│  └──────────────────────────────────┘        │
│                                              │
│        [ 开始比对 ]                            │
│                                              │
│  ┌─ 比对日志 ──────────────────────┐         │
│  │ 志愿 1: [A246] 复旦大学           │        │
│  │   [未录取] 150 > 248             │        │
│  │ 志愿 2: ...                      │        │
│  │ 志愿 7: [A246] 复旦大学           │        │
│  │   [录取] 150 <= 583             │        │
│  └──────────────────────────────────┘        │
└──────────────────────────────────────────────┘
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

## 比对逻辑详解

```
对于志愿表中的每一行：
  步骤1：提取 院校代码（如 A246）和 专业代码（如 33）
  步骤2：在投档线中查找 (院校代码, 专业代码) 组合
  步骤3：如果找到：
           你的位次 ≤ 最低位次 → 录取！停止比对
           你的位次 > 最低位次 → 不录取，继续下一条
  步骤4：如果未找到：
           提示"代码组合未找到"（可能是专业代码跨年变化），继续下一条
```

- **位次**：全省排名，数字越小越好
- **最低位次**：去年该专业录取的最后一名学生的位次
- **判断规则**：你的位次 ≤ 最低位次 → 可能被录取

## 项目结构

```
志愿表-投档线对照/
├── gui.py                    ← 图形界面主程序（Material Design 3）
├── compare.py                ← 命令行版本
├── cloud.ico                 ← 应用图标
├── cloud.png                 ← 应用图标（PNG 备用）
├── 志愿-投档线对照.spec       ← PyInstaller 打包配置文件
├── BUILD.md                  ← 构建与打包详细指南
├── README.md                 ← 英文 README
├── README.zh.md              ← 本文件（中文 README）
├── 志愿表.pdf                 ← （用户提供）考生志愿表
├── 投档线.xls                 ← （用户提供）投档线数据
└── comparison_log_*.txt      ← （自动生成）比对日志
```

## 两种版本

| 版本 | 文件 | 适用场景 |
|------|------|---------|
| **图形界面** | `gui.py` | 日常使用，拖拽操作，弹窗提示 |
| **命令行** | `compare.py` | 脚本化/自动化，`python compare.py` 直接运行 |

## 打包为独立 EXE（无需 Python 环境）

### 前置准备

```bash
pip install pyinstaller pillow
```

### 打包命令

在项目目录下执行：

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

> **PowerShell 用户**：把 `\` 换成 `` ` `` 换行。

### 参数解释

| 参数 | 作用 |
|------|------|
| `--name` | 输出 exe 文件名 |
| `--windowed` | 不弹命令行黑窗（纯 GUI 模式） |
| `--onefile` | 打包成单个 exe，方便分发 |
| `--icon cloud.ico` | exe 文件图标 |
| `--add-data "cloud.ico;."` | 把图标文件一同打包 |
| `--hidden-import` | 显式声明 PyInstaller 检测不到的隐式导入 |
| `--collect-all tkinterdnd2` | 收集拖拽库的 Tcl/Tk 扩展 dll |

### 输出位置

```
dist/
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
--add-data "C:\Users\...\Python311\tcl\tkdnd2.8;tkdnd"
```

**Q: 杀毒软件报毒**

单文件 exe 容易触发误报，改用 `--onedir` 可大幅降低误报率。

## 故障排除

| 问题 | 解决方法 |
|------|---------|
| 启动报错 `No module named 'tkinter'` | Linux 需安装 `python3-tk` |
| 拖拽文件没反应 | 点击文件区域手动浏览选择，或检查 `tkinterdnd2` 是否安装 |
| `UnicodeEncodeError` 编码错误 | 终端执行前设置 `set PYTHONIOENCODING=utf-8` |
| 日志中文显示乱码 | 检查系统是否安装了中文字体（如 Microsoft YaHei） |

## 数据格式说明

### 志愿表 PDF

程序提取以下信息：
- 志愿号
- 院校代码 + 院校名称（如 `A246复旦大学`）
- 专业代码 + 专业名称（如 `33计算机科学与技术`）
- 办学性质（公办/民办/独立学院 等）

### 投档线 XLS

Excel 文件应包含以下列：
- **专业代码及名称** — 如 `33计算机科学与技术`
- **院校代码及名称** — 如 `A246复旦大学`
- **投档计划数** — 该专业的计划招生人数
- **最低位次** — 该专业录取的最后一名学生的位次

## 许可

[MIT](LICENSE)

---

使用 Python + tkinter + Material Design 3 构建。
