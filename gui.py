#!/usr/bin/env python3
"""
志愿-投档线对照分析 — 图形化界面 (Material Design 3)
====================================================
拖入志愿表PDF和投档线XLS文件，输入位次，一键比对。
比对过程输出到日志文件，录取结果弹窗提示。

用法: python gui.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont
import threading
import re
import os
import sys
import ctypes
import unicodedata
from datetime import datetime

# -- 高DPI适配 -----------------------------------------
# 在创建任何 tkinter 窗口之前必须调用，否则字体在高分屏模糊
if sys.platform == 'win32':
    try:
        # Windows 10 1703+: Per-Monitor DPI V2 (最佳清晰度)
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            # Windows 8.1+
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                # Windows Vista/7
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

# ── 尝试导入拖拽支持 ─────────────────────────────────
try:
    from tkinterdnd2 import TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

# ── 业务库（延迟导入以避免启动慢） ───────────────────
pd = None
PyPDF2 = None

def _ensure_deps():
    global pd, PyPDF2
    if pd is None:
        import pandas as pd_impl
        pd = pd_impl
    if PyPDF2 is None:
        import PyPDF2 as pdf_impl
        PyPDF2 = pdf_impl


# ═══════════════════════════════════════════════════════
#  Material Design 3 色彩系统
# ═══════════════════════════════════════════════════════

class MD3:
    """MD3 色彩令牌 & 字体"""
    # Primary
    PRIMARY        = "#6750A4"
    ON_PRIMARY     = "#FFFFFF"
    PRIMARY_CONT   = "#EADDFF"
    ON_PRIMARY_CONT = "#21005D"
    # Surface
    SURFACE        = "#FFFBFE"
    SURFACE_LOW    = "#F7F2FA"
    SURFACE_HIGH   = "#ECE6F0"
    ON_SURFACE     = "#1C1B1F"
    ON_SURFACE_V   = "#49454F"
    # Outline
    OUTLINE        = "#79747E"
    OUTLINE_V      = "#CAC4D0"
    # Error
    ERROR          = "#B3261E"
    ON_ERROR       = "#FFFFFF"
    ERROR_CONT     = "#F9DEDC"
    # Success
    SUCCESS        = "#1B7A2E"
    SUCCESS_CONT   = "#D8F3DC"
    # Warning
    WARN           = "#8A6100"
    WARN_CONT      = "#FFF2C5"

    # Elevation colors (simulated with borders)
    SHADOW_LIGHT   = "#E0E0E0"
    SHADOW_DARK    = "#BDBDBD"

    @staticmethod
    def fonts(root):
        """返回预定义字体"""
        return {
            'display': tkfont.Font(root=root, family='Microsoft YaHei UI', size=22, weight='bold'),
            'headline': tkfont.Font(root=root, family='Microsoft YaHei UI', size=16, weight='bold'),
            'title':    tkfont.Font(root=root, family='Microsoft YaHei UI', size=13, weight='bold'),
            'body':     tkfont.Font(root=root, family='Microsoft YaHei UI', size=11),
            'label':    tkfont.Font(root=root, family='Microsoft YaHei UI', size=10),
            'caption':  tkfont.Font(root=root, family='Microsoft YaHei UI', size=9),
            'mono':     tkfont.Font(root=root, family='Cascadia Code', size=10),
        }


# ═══════════════════════════════════════════════════════
#  MD3 风格组件
# ═══════════════════════════════════════════════════════

class MD3Card(tk.Frame):
    """MD3 卡片容器 — 带圆角和阴影的Frame"""
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=MD3.SURFACE, **kw)
        # 模拟阴影边框
        self.shadow = tk.Frame(parent, bg=MD3.SHADOW_LIGHT)
        self.inner = tk.Frame(self, bg=MD3.SURFACE)
        self.inner.pack(fill='both', expand=True, padx=1, pady=1)

    def pack(self, **kw):
        self.shadow.pack(**kw)
        super().pack(in_=self.shadow, fill='both', expand=True, padx=0, pady=0)

    def grid(self, **kw):
        self.shadow.grid(**kw)
        super().grid(in_=self.shadow, fill='both', expand=True, padx=0, pady=0)


class MD3DropZone(tk.Frame):
    """MD3 拖拽/点击选择文件区域"""
    def __init__(self, parent, label, file_types, on_file, fonts, **kw):
        super().__init__(parent, bg=MD3.SURFACE_HIGH, cursor='hand2', **kw)
        self.label_text = label
        self.file_types = file_types
        self.on_file = on_file
        self.fonts = fonts
        self.file_path = None

        # 内容
        self.icon_label = tk.Label(self, text="📂", font=fonts['display'],
                                   bg=MD3.SURFACE_HIGH, fg=MD3.PRIMARY)
        self.icon_label.pack(pady=(20, 8))

        self.text_label = tk.Label(self, text=label, font=fonts['title'],
                                   bg=MD3.SURFACE_HIGH, fg=MD3.ON_SURFACE,
                                   wraplength=220)
        self.text_label.pack(pady=(0, 4))

        self.hint_label = tk.Label(self, text="拖拽文件到此处，或点击选择",
                                   font=fonts['caption'],
                                   bg=MD3.SURFACE_HIGH, fg=MD3.ON_SURFACE_V)
        self.hint_label.pack(pady=(0, 4))

        self.status_label = tk.Label(self, text="尚未选择文件",
                                     font=fonts['label'],
                                     bg=MD3.SURFACE_HIGH, fg=MD3.ON_SURFACE_V)
        self.status_label.pack(pady=(0, 16))

        # 点击事件
        self.bind("<Button-1>", self._on_click)
        self.icon_label.bind("<Button-1>", self._on_click)
        self.text_label.bind("<Button-1>", self._on_click)
        self.hint_label.bind("<Button-1>", self._on_click)
        self.status_label.bind("<Button-1>", self._on_click)

        # Hover效果
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

        # 拖拽支持
        if DND_AVAILABLE:
            try:
                self.drop_target_register('DND_Files')
                self.dnd_bind('<<Drop>>', self._on_drop)
            except Exception:
                pass

    def _on_click(self, event):
        path = filedialog.askopenfilename(
            title=f"选择{self.label_text}",
            filetypes=[(f"{self.label_text}文件", f"*.{self.file_types[0]}"),
                       ("所有文件", "*.*")]
        )
        if path:
            self.set_file(path)

    def _on_drop(self, event):
        # tkinterdnd2 返回花括号包裹的路径
        data = event.data.strip()
        if data.startswith('{') and data.endswith('}'):
            data = data[1:-1]
        # 可能包含多个文件，取第一个
        paths = data.split('} {') if '} {' in data else [data]
        path = paths[0].strip('{}')
        if path and any(path.lower().endswith(f'.{ext}') for ext in self.file_types):
            self.set_file(path)

    def _on_enter(self, event):
        self.configure(bg=MD3.PRIMARY_CONT)
        for child in (self.icon_label, self.text_label, self.hint_label, self.status_label):
            child.configure(bg=MD3.PRIMARY_CONT)

    def _on_leave(self, event):
        self.configure(bg=MD3.SURFACE_HIGH)
        for child in (self.icon_label, self.text_label, self.hint_label, self.status_label):
            child.configure(bg=MD3.SURFACE_HIGH)

    def set_file(self, path):
        self.file_path = path
        fname = os.path.basename(path)
        self.status_label.configure(text=f"✓ {fname}", fg=MD3.SUCCESS)
        self.on_file(path)

    def reset(self):
        self.file_path = None
        self.status_label.configure(text="尚未选择文件", fg=MD3.ON_SURFACE_V)


class MD3Button(tk.Frame):
    """MD3 Filled Button"""
    def __init__(self, parent, text, command, fonts, **kw):
        super().__init__(parent, bg=MD3.SURFACE, **kw)
        self.command = command
        self.enabled = True

        self.btn = tk.Label(self, text=text, font=fonts['title'],
                            bg=MD3.PRIMARY, fg=MD3.ON_PRIMARY,
                            padx=32, pady=10, cursor='hand2')
        self.btn.pack()

        self.btn.bind("<Button-1>", self._on_click)
        self.btn.bind("<Enter>", lambda e: self.btn.configure(bg="#7A67B8") if self.enabled else None)
        self.btn.bind("<Leave>", lambda e: self.btn.configure(bg=MD3.PRIMARY) if self.enabled else None)

    def _on_click(self, event):
        if self.enabled and self.command:
            self.command()

    def set_enabled(self, state):
        self.enabled = state
        if state:
            self.btn.configure(bg=MD3.PRIMARY, cursor='hand2')
        else:
            self.btn.configure(bg=MD3.OUTLINE_V, cursor='arrow')


class MD3TextField(tk.Frame):
    """MD3 Outlined Text Field"""
    def __init__(self, parent, label, fonts, **kw):
        super().__init__(parent, bg=MD3.SURFACE, **kw)

        self.var = tk.StringVar()
        self.fonts = fonts

        # 标签
        self.label = tk.Label(self, text=label, font=fonts['label'],
                              bg=MD3.SURFACE, fg=MD3.PRIMARY)
        self.label.pack(anchor='w')

        # 输入框（模拟 outlined 样式）
        self.entry_frame = tk.Frame(self, bg=MD3.OUTLINE, padx=2, pady=2)
        self.entry_frame.pack(fill='x')

        self.entry = tk.Entry(self.entry_frame, textvariable=self.var,
                              font=fonts['body'], bg=MD3.SURFACE,
                              fg=MD3.ON_SURFACE, relief='flat',
                              insertbackground=MD3.PRIMARY,
                              insertwidth=2)
        self.entry.pack(fill='x', ipady=6, padx=8)

    def get(self):
        return self.var.get().strip()

    def set(self, value):
        self.var.set(value)


# ═══════════════════════════════════════════════════════
#  核心比对逻辑（复用 compare.py）
# ═══════════════════════════════════════════════════════

def sanitize(text):
    """将CJK兼容字符转为普通中文，避免编码错误"""
    text = unicodedata.normalize('NFKC', text)
    result = []
    for ch in text:
        try:
            ch.encode('gbk')
            result.append(ch)
        except UnicodeEncodeError:
            try:
                decomposed = unicodedata.normalize('NFKD', ch)
                for dc in decomposed:
                    try:
                        dc.encode('gbk')
                        result.append(dc)
                        break
                    except UnicodeEncodeError:
                        continue
                else:
                    result.append('?')
            except Exception:
                result.append('?')
    return ''.join(result)


def extract_pdf_text(pdf_path):
    reader = PyPDF2.PdfReader(pdf_path)
    full_text = ""
    for page in reader.pages:
        text = page.extract_text()
        if text:
            full_text += text + "\n"
    return full_text


def parse_volunteer_table(text):
    lines = text.split('\n')
    start_idx = 0
    for i, line in enumerate(lines):
        if '志愿号' in line and '报考院校' in line:
            start_idx = i + 1
            break

    entry_pattern = re.compile(r'^(\d+)\s*([A-Z]\d{3})')
    raw_entries = []
    current_lines = []

    for i in range(start_idx, len(lines)):
        line = lines[i]
        match = entry_pattern.match(line)
        if match:
            if current_lines:
                raw_entries.append(''.join(current_lines))
            current_lines = [line]
        else:
            if current_lines:
                current_lines.append(line)

    if current_lines:
        raw_entries.append(''.join(current_lines))

    volunteers = []
    for entry_text in raw_entries:
        vol = _parse_single(entry_text)
        if vol:
            volunteers.append(vol)
    return volunteers


def _parse_single(text):
    if '查看时间' in text:
        text = text.split('查看时间')[0]

    text = text.strip()
    if not text:
        return None

    vol_match = re.match(r'^(\d+)\s*', text)
    if not vol_match:
        return None

    school_match = re.search(r'([A-Z]\d{3})', text)
    if not school_match:
        return None

    school_code = school_match.group(1)
    school_start = school_match.end()
    nature_match = re.search(r'(公办院校|民办院校|独立学院|中外合作办学|内地与港澳台地区合作办学)', text)

    if nature_match:
        school_name = text[school_start:nature_match.start()].strip()
        after_nature = text[nature_match.end():]
    else:
        school_name = text[school_start:].strip()
        after_nature = text

    major_match = re.match(r'\s*([0-9A-Za-z]+)\s*', after_nature)
    if not major_match:
        return None
    major_code = major_match.group(1)
    rest = after_nature[major_match.end():].strip()

    level_match = re.search(r'(本科|专科|高职)', rest)
    if level_match:
        full_major_text = rest[:level_match.start()].strip()
    else:
        full_major_text = rest.strip()

    short_name = re.split(r'[（(]', full_major_text)[0].strip()
    short_name = re.sub(r'\s+', '', short_name)

    return {
        '志愿号': vol_match.group(0).strip(),
        '院校代码': school_code,
        '院校名称': sanitize(school_name),
        '专业代码': major_code,
        '专业名称': sanitize(short_name),
        '专业全称': sanitize(full_major_text),
    }


def parse_toudang_xls(xls_path):
    df = pd.read_excel(xls_path, header=None, skiprows=1)
    df.columns = ['专业代码及名称', '院校代码及名称', '投档计划数', '最低位次']
    lookup = {}
    for _, row in df.iterrows():
        major_full = str(row['专业代码及名称']).strip()
        school_full = str(row['院校代码及名称']).strip()
        plan = str(row['投档计划数']).strip()
        min_rank_raw = str(row['最低位次']).strip()

        school_match = re.match(r'^([A-Z]\d{3})', school_full)
        major_match = re.match(r'^([0-9A-Za-z]+)', major_full)
        if not school_match or not major_match:
            continue

        school_code = school_match.group(1)
        major_code = major_match.group(1)
        major_name = major_full[major_match.end():]

        try:
            min_rank = int(min_rank_raw)
        except ValueError:
            rank_clean = re.sub(r'[^\d]', '', min_rank_raw)
            min_rank = int(rank_clean) if rank_clean else None

        lookup[(school_code, major_code)] = {
            '院校名称': sanitize(school_full[school_match.end():]),
            '专业名称': sanitize(major_name),
            '投档计划数': plan,
            '最低位次': min_rank,
            '最低位次原始': min_rank_raw,
        }
    return lookup


# ═══════════════════════════════════════════════════════
#  主 GUI 应用程序
# ═══════════════════════════════════════════════════════

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("志愿-投档线对照分析")
        self.root.geometry("800x720")
        self.root.configure(bg=MD3.SURFACE)
        self.root.minsize(640, 600)

        self.fonts = MD3.fonts(root)

        # 应用状态
        self.pdf_path = None
        self.xls_path = None
        self.is_running = False
        self.log_lines = []

        self._build_ui()

    def _build_ui(self):
        # ── 标题栏 ──
        title_bar = tk.Frame(self.root, bg=MD3.SURFACE, pady=12)
        title_bar.pack(fill='x', padx=24)

        tk.Label(title_bar, text="志愿-投档线对照分析",
                 font=self.fonts['display'], bg=MD3.SURFACE,
                 fg=MD3.ON_SURFACE).pack(anchor='w')

        # ── 分隔线 ──
        sep = tk.Frame(self.root, bg=MD3.OUTLINE_V, height=1)
        sep.pack(fill='x', padx=24)

        # ── 文件选择区域 ──
        file_area = tk.Frame(self.root, bg=MD3.SURFACE)
        file_area.pack(fill='x', padx=24, pady=(16, 8))

        # 左右两列
        file_area.columnconfigure(0, weight=1)
        file_area.columnconfigure(1, weight=1)

        self.drop_pdf = MD3DropZone(file_area, "志愿表 PDF",
                                     file_types=['pdf'],
                                     on_file=self._on_pdf_selected,
                                     fonts=self.fonts)
        self.drop_pdf.grid(row=0, column=0, sticky='nsew', padx=(0, 8))

        self.drop_xls = MD3DropZone(file_area, "投档线 XLS",
                                     file_types=['xls', 'xlsx'],
                                     on_file=self._on_xls_selected,
                                     fonts=self.fonts)
        self.drop_xls.grid(row=0, column=1, sticky='nsew', padx=(8, 0))

        # ── 位次输入区域 ──
        input_area = tk.Frame(self.root, bg=MD3.SURFACE)
        input_area.pack(fill='x', padx=24, pady=(8, 12))

        rank_card = tk.Frame(input_area, bg=MD3.SURFACE_HIGH, padx=16, pady=12)
        rank_card.pack(fill='x')

        tk.Label(rank_card, text="输入你的位次",
                 font=self.fonts['title'], bg=MD3.SURFACE_HIGH,
                 fg=MD3.ON_SURFACE).pack(anchor='w')

        self.rank_field = MD3TextField(rank_card, "位次（整数，越小越靠前）", self.fonts)
        self.rank_field.pack(fill='x', pady=(4, 4))

        # ── 开始按钮 ──
        btn_area = tk.Frame(self.root, bg=MD3.SURFACE)
        btn_area.pack(fill='x', padx=24, pady=(0, 12))

        self.start_btn = MD3Button(btn_area, "开始比对", self._start_comparison, self.fonts)
        self.start_btn.pack()

        # 进度指示
        self.progress = ttk.Progressbar(btn_area, mode='indeterminate')
        self.progress_label = tk.Label(btn_area, text="", font=self.fonts['caption'],
                                       bg=MD3.SURFACE, fg=MD3.ON_SURFACE_V)

        # ── 日志输出区域 ──
        log_label = tk.Frame(self.root, bg=MD3.SURFACE)
        log_label.pack(fill='x', padx=24, pady=(0, 4))
        tk.Label(log_label, text="比对日志", font=self.fonts['title'],
                 bg=MD3.SURFACE, fg=MD3.ON_SURFACE).pack(side='left')
        self.log_status = tk.Label(log_label, text="等待开始...", font=self.fonts['caption'],
                                   bg=MD3.SURFACE, fg=MD3.ON_SURFACE_V)
        self.log_status.pack(side='right')

        # 日志文本框
        log_frame = tk.Frame(self.root, bg=MD3.OUTLINE, padx=2, pady=2)
        log_frame.pack(fill='both', expand=True, padx=24, pady=(0, 16))

        self.log_text = tk.Text(log_frame, font=self.fonts['mono'],
                                bg='#1E1E1E', fg='#D4D4D4',
                                insertbackground='white',
                                relief='flat', wrap='word',
                                padx=12, pady=8)
        self.log_text.pack(fill='both', expand=True)

        # 配置日志颜色标签
        self.log_text.tag_configure('info', foreground='#8FC7FF')
        self.log_text.tag_configure('success', foreground='#6A9955')
        self.log_text.tag_configure('error', foreground='#F44747')
        self.log_text.tag_configure('warn', foreground='#CE9178')
        self.log_text.tag_configure('header', foreground='#DCDCAA')
        self.log_text.tag_configure('dim', foreground='#808080')
        self.log_text.configure(state='disabled')

    # ── 事件处理 ──

    def _on_pdf_selected(self, path):
        self.pdf_path = path
        self._log(f"[INFO] 志愿表PDF已加载: {os.path.basename(path)}")

    def _on_xls_selected(self, path):
        self.xls_path = path
        self._log(f"[INFO] 投档线XLS已加载: {os.path.basename(path)}")

    def _log(self, msg, tag='info'):
        """写入日志到GUI和内存"""
        self.log_lines.append(msg)
        self.log_text.configure(state='normal')
        self.log_text.insert('end', msg + '\n', tag)
        self.log_text.see('end')
        self.log_text.configure(state='disabled')

    def _start_comparison(self):
        """开始比对（在新线程中运行）"""
        # 验证
        if self.is_running:
            return

        if not self.pdf_path:
            messagebox.showwarning("提示", "请先选择志愿表PDF文件。", parent=self.root)
            return
        if not self.xls_path:
            messagebox.showwarning("提示", "请先选择投档线XLS文件。", parent=self.root)
            return

        rank_str = self.rank_field.get()
        if not rank_str:
            messagebox.showwarning("提示", "请输入你的位次。", parent=self.root)
            return
        try:
            user_rank = int(rank_str)
            if user_rank <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("提示", "位次必须为正整数。", parent=self.root)
            return

        # 开始
        self.is_running = True
        self.start_btn.set_enabled(False)
        self.log_status.configure(text="正在比对...")
        self.progress.pack(pady=(4, 0))
        self.progress.start(10)
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.configure(state='disabled')
        self.log_lines.clear()

        thread = threading.Thread(target=self._run_comparison, args=(user_rank,), daemon=True)
        thread.start()

    def _run_comparison(self, user_rank):
        """在后台线程中执行比对"""
        try:
            _ensure_deps()

            self._log("=" * 60, 'header')
            self._log("  志愿-投档线对照分析", 'header')
            self._log(f"  开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", 'dim')
            self._log(f"  你的位次: {user_rank}", 'header')
            self._log("=" * 60, 'header')

            # 解析志愿表
            self._log("\n[STEP 1/3] 解析志愿表PDF...", 'info')
            pdf_text = extract_pdf_text(self.pdf_path)

            info_match = re.search(r'姓名[：:]\s*(\S+)', pdf_text)
            score_match = re.search(r'成绩[：:]\s*(\d+)\s*(\d+)', pdf_text)
            student_name = sanitize(info_match.group(1)) if info_match else "未知"
            student_score = f"{score_match.group(1)}{score_match.group(2)}" if score_match else "未知"
            self._log(f"  考生: {student_name}  成绩: {student_score}", 'dim')

            volunteers = parse_volunteer_table(pdf_text)
            self._log(f"  共解析到 {len(volunteers)} 条志愿", 'info')

            # 解析投档线
            self._log("\n[STEP 2/3] 解析投档线XLS...", 'info')
            lookup = parse_toudang_xls(self.xls_path)
            self._log(f"  共解析到 {len(lookup)} 条投档记录", 'info')

            # 逐条比对
            self._log(f"\n[STEP 3/3] 逐条比对（共{len(volunteers)}条）...\n", 'info')

            admitted = False
            not_found_count = 0
            checked_count = 0

            for vol in volunteers:
                checked_count += 1
                vn = vol['志愿号']
                sc = vol['院校代码']
                sn = vol['院校名称']
                mc = vol['专业代码']
                mn = vol['专业名称']

                self._log(f"--- 志愿 {vn} ---", 'dim')
                self._log(f"  [{sc}] {sn}  |  专业代码: {mc}  {mn}")

                key = (sc, mc)
                if key not in lookup:
                    self._log(f"  [WARN] 投档线中未找到该代码组合", 'warn')
                    not_found_count += 1
                    continue

                info = lookup[key]
                min_rank = info['最低位次']

                self._log(f"  投档专业: {info['专业名称']}", 'dim')
                self._log(f"  投档计划数: {info['投档计划数']}", 'dim')
                self._log(f"  最低位次: {info['最低位次原始']}", 'dim')

                if min_rank is None:
                    self._log(f"  [WARN] 最低位次无法解析", 'warn')
                    continue

                if user_rank <= min_rank:
                    self._log(f"  [录取] 位次{user_rank} <= 最低位次{min_rank}", 'success')
                    admitted = True

                    result_msg = (
                        f"在第 {vn} 志愿被录取!\n\n"
                        f"院校: [{sc}] {sn}\n"
                        f"专业: [{mc}] {mn}\n"
                        f"投档专业: {info['专业名称']}\n"
                        f"你的位次: {user_rank}\n"
                        f"最低位次: {min_rank}\n"
                        f"投档计划数: {info['投档计划数']}"
                    )
                    self.root.after(0, lambda: self._show_result(True, result_msg))
                    break
                else:
                    self._log(f"  [未录取] 位次{user_rank} > 最低位次{min_rank}", 'error')

            if not admitted:
                if not_found_count == len(volunteers):
                    result_msg = "所有志愿在投档线中均未找到匹配记录。\n(可能是专业代码跨年变化导致)"
                else:
                    result_msg = f"遍历全部 {len(volunteers)} 个志愿，均未达到投档线。"

                self._log(f"\n  {result_msg}", 'warn')
                self.root.after(0, lambda: self._show_result(False, result_msg))

            self._log(f"\n检查了 {checked_count}/{len(volunteers)} 条志愿", 'dim')
            self._log(f"其中 {not_found_count} 条未在投档线中找到匹配", 'dim')

        except Exception as e:
            import traceback
            err = traceback.format_exc()
            self._log(f"\n[ERROR] 比对过程出错:\n{err}", 'error')
            self.root.after(0, lambda: messagebox.showerror("错误", f"比对过程出错:\n{str(e)}", parent=self.root))
        finally:
            self.root.after(0, self._comparison_done)

    def _comparison_done(self):
        """比对完成后的清理"""
        self.is_running = False
        self.start_btn.set_enabled(True)
        self.progress.stop()
        self.progress.pack_forget()
        self.log_status.configure(text="比对完成")
        self._save_log_file()

    def _show_result(self, admitted, msg):
        """弹窗显示录取结果"""
        if admitted:
            messagebox.showinfo("录取结果", msg, parent=self.root)
        else:
            messagebox.showwarning("录取结果", msg, parent=self.root)

    def _save_log_file(self):
        """保存日志文件"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        log_path = f"comparison_log_{timestamp}.txt"
        try:
            with open(log_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(self.log_lines))
            self._log(f"\n日志已保存至: {log_path}", 'dim')
            self.log_status.configure(text=f"日志已保存: {log_path}")
        except Exception as e:
            self._log(f"\n[WARN] 日志保存失败: {e}", 'warn')


# ═══════════════════════════════════════════════════════
#  入口
# ═══════════════════════════════════════════════════════

def _get_dpi_scale(root):
    """获取系统 DPI 缩放因子，用于字体大小调整"""
    try:
        # 获取真实 DPI（Windows 10+）
        import ctypes
        hdc = ctypes.windll.user32.GetDC(0)
        dpi_x = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
        ctypes.windll.user32.ReleaseDC(0, hdc)
        # 标准 DPI = 96，缩放因子 = 实际 DPI / 96
        return dpi_x / 96.0
    except Exception:
        return 1.0


def _get_icon_path():
    """获取图标文件路径（兼容 PyInstaller 打包后的路径）"""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    # 优先 ico（用于 Windows 任务栏），其次 png
    ico = os.path.join(base, 'cloud.ico')
    png = os.path.join(base, 'cloud.png')
    if os.path.exists(ico):
        return ico
    elif os.path.exists(png):
        return png
    return None


def main():
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    # 根据系统 DPI 缩放 tk 内部比例（消除模糊的关键）
    scale = _get_dpi_scale(root)
    if scale > 1.0:
        root.tk.call('tk', 'scaling', scale)

    # 设置窗口图标
    icon_path = _get_icon_path()
    if icon_path:
        try:
            if icon_path.endswith('.ico'):
                root.iconbitmap(icon_path)
            else:
                icon = tk.PhotoImage(file=icon_path)
                root.iconphoto(True, icon)
                root._icon = icon  # 保持引用防止被 GC
        except Exception:
            pass  # 图标加载失败不影响程序运行

    app = App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
