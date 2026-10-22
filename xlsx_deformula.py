#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 公式转数值工具 (Excel Formula to Value)

功能
----
1) 图形界面选择要处理的 Excel 文件（可多选，支持 .xlsx / .xlsm）
2) 指定输出位置（源文件同目录自动命名 / 指定目录 / 指定单个文件名）
3) 原始文件全程只读，绝不修改（只读取源文件，结果写入新文件）
4) 保留图表、图片、合并单元格、条件格式、行高列宽——只在 XML 层剥离 <f> 公式，
   保留公式的缓存计算结果 <v>，不用 openpyxl 重新保存（会丢图表）

零第三方依赖：仅用 Python 标准库（tkinter + zipfile + re）。
可作为命令行工具使用：  python xlsx_deformula.py --cli 源文件.xlsx [输出.xlsx]

注意：本工具不计算公式。若原文件从未被 Excel/WPS 保存过（公式没有缓存结果），
这些单元格处理后会变成空值——程序会在处理前统计并提示。
"""

import os
import re
import sys
import time
import shutil
import zipfile
import threading
import traceback
from xml.sax.saxutils import escape

APP_NAME = "Excel 公式转数值工具"
APP_VER = "1.0"
SUFFIX = "_无公式"

# ---------------------------------------------------------------- 核心引擎

CELL_RE = re.compile(r'<c\b[^>]*?(?:/>|>.*?</c>)', re.S)
F_RE = re.compile(r'<f\b[^>]*?/>|<f\b[^>]*?>.*?</f>', re.S)
ATTR_RE = re.compile(r'([\w:.-]+)="([^"]*)"')
DISPIMG_RE = re.compile(r'(?:_xlfn\.)?DISPIMG\("([^"]+)"', re.I)


def _unescape(s):
    return (s.replace('&lt;', '<').replace('&gt;', '>')
             .replace('&quot;', '"').replace('&apos;', "'")
             .replace('&amp;', '&'))


def _strip_f(block):
    """去掉 <f>，保留 <v>。公式结果是字符串(t=str)时转 inlineStr，否则文本非法。

    特例：WPS 的内嵌图片单元格 =DISPIMG("ID",1) —— 它不是计算型公式而是图片载体，
    原样保留（公式与格式都不动），否则图片会消失。
    返回 (新XML块, 是否剥离, 是否文本转换, 是否跳过) 。
    """
    if '<f' not in block:
        return block, 0, 0, False
    if DISPIMG_RE.search(_unescape(block)):
        return block, 0, 0, True
    new = F_RE.sub('', block)
    open_end = new.index('>')
    attrs = ATTR_RE.findall(new[:open_end + 1])
    adict = dict(attrs)
    t = adict.get('t')
    v_m = re.search(r'<v>(.*?)</v>', new, re.S)
    v = v_m.group(1) if v_m else None
    base = ' '.join('%s="%s"' % (k, val) for k, val in attrs if k != 't')
    if t == 'str':
        if v:
            txt = escape(_unescape(v))
            return ('<c %s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                    % (base, txt)), 1, 1, None
        return '<c %s/>' % base, 1, 1, None
    if t == 'e':
        if v:
            return '<c %s t="e"><v>%s</v></c>' % (base, v), 1, 0, None
        return '<c %s/>' % base, 1, 0, None
    return new, 1, 0, None


def scan_workbook(path):
    """处理前预检：公式单元格数、无缓存值的公式数、内嵌图片单元格数。"""
    total = missing = keep_image = 0
    sheets = 0
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.startswith('xl/worksheets/') and n.endswith('.xml'):
                sheets += 1
                xml = z.read(n).decode('utf-8', 'replace')
                for blk in CELL_RE.findall(xml):
                    if '<f' not in blk:
                        continue
                    if DISPIMG_RE.search(_unescape(blk)):
                        keep_image += 1
                        continue
                    total += 1
                    if '<v>' not in blk or re.search(r'<v></v>|<v/>', blk):
                        missing += 1
    return {'sheets': sheets, 'formulas': total, 'no_cache': missing,
            'keep_image': keep_image}


def convert(src, dst, progress=None, keep_external=True):
    """把 src 的所有公式转成数值写入 dst（src 只读）。返回统计 dict。"""
    src = os.path.abspath(src)
    dst = os.path.abspath(dst)
    if src == dst:
        raise ValueError("输出文件不能与原文件相同（原始文件必须保持不变）")
    t0 = time.time()

    def say(msg, frac=None):
        if progress:
            progress(msg, frac)

    say("正在读取源文件…", 0.02)
    pre = scan_workbook(src)
    stats = {'sheets': 0, 'formulas': 0, 'text': 0, 'keep_image': 0,
             'no_cache': pre['no_cache'], 'seconds': 0.0}

    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        sheet_xml = {}
        for part in [n for n in names if n.startswith('xl/worksheets/') and n.endswith('.xml')]:
            xml = z.read(part).decode('utf-8', 'replace')
            counter = {'f': 0, 'img': 0}

            def repl(m):
                blk, removed, conv, skipped = _strip_f(m.group(0))
                if skipped:
                    counter['img'] += 1
                else:
                    counter['f'] += removed
                    stats['text'] += conv
                return blk

            sheet_xml[part] = CELL_RE.sub(repl, xml)
            stats['formulas'] += counter['f']
            stats['keep_image'] += counter['img']
            if counter['f']:
                stats['sheets'] += 1

    say("正在写入新文件…", 0.35)
    tmp = dst + '.writing'
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        drop = {'xl/calcChain.xml'}
        if not keep_external:
            drop |= {n for n in zin.namelist()
                     if n.startswith('xl/externalLinks/')}
        total = len(zin.infolist())
        for i, item in enumerate(zin.infolist()):
            n = item.filename
            if n in drop:
                continue
            data = zin.read(n)
            if n in sheet_xml:
                data = sheet_xml[n].encode('utf-8')
            elif n == '[Content_Types].xml':
                s = data.decode('utf-8', 'replace')
                s = re.sub(r'<Override[^>]*calcChain\.xml"[^>]*/>', '', s)
                if not keep_external:
                    s = re.sub(r'<Override[^>]*externalLink\d*\.xml"[^>]*/>', '', s)
                data = s.encode('utf-8')
            elif n == 'xl/_rels/workbook.xml.rels':
                s = data.decode('utf-8', 'replace')
                s = re.sub(r'<Relationship[^>]*calcChain\.xml"[^>]*/>', '', s)
                if not keep_external:
                    s = re.sub(r'<Relationship[^>]*externalLink\d*\.xml"[^>]*/>', '', s)
                data = s.encode('utf-8')
            elif n == 'xl/workbook.xml':
                s = data.decode('utf-8', 'replace')
                s = s.replace(' fullCalcOnLoad="1"', '')
                if not keep_external:
                    s = re.sub(r'<externalReferences>.*?</externalReferences>', '', s, flags=re.S)
                data = s.encode('utf-8')
            zout.writestr(item, data)
            if progress and i % 20 == 0:
                say("正在写入新文件… (%d/%d)" % (i, total), 0.35 + 0.5 * i / max(total, 1))
    shutil.move(tmp, dst)

    say("正在自检…", 0.9)
    left = 0
    with zipfile.ZipFile(dst) as z:
        bad = z.testzip()
        for n in z.namelist():
            if n.startswith('xl/worksheets/') and n.endswith('.xml'):
                sx = z.read(n).decode('utf-8', 'replace')
                for blk in CELL_RE.findall(sx):
                    # 内嵌图片单元格的 DISPIMG 公式是刻意保留的，不计入残留
                    if '<f' in blk and not DISPIMG_RE.search(_unescape(blk)):
                        left += 1
    if bad:
        raise IOError("写入校验失败（zip 损坏）：%s" % bad)
    if left:
        raise IOError("仍有 %d 个公式未被移除，处理中止" % left)
    stats['seconds'] = round(time.time() - t0, 1)
    say("完成", 1.0)
    return stats


# ---------------------------------------------------------------- 命令行模式

def run_cli(argv):
    if not argv:
        print(__doc__)
        return 1
    src = argv[0]
    dst = argv[1] if len(argv) > 1 else default_out(src, None)
    pre = scan_workbook(src)
    print("源文件 : %s" % src)
    print("工作表 : %d   公式单元格: %d   其中无缓存值: %d   内嵌图片单元格(保留不动): %d"
          % (pre['sheets'], pre['formulas'], pre['no_cache'], pre['keep_image']))
    st = convert(src, dst, progress=lambda m, f=None: print("  " + m))
    print("输出   : %s" % dst)
    print("结果   : 剥离公式 %d 个 / 文本转换 %d / 保留图片单元格 %d 个 / 耗时 %.1fs"
          % (st['formulas'], st['text'], st['keep_image'], st['seconds']))
    return 0


def default_out(src, outdir):
    base = os.path.splitext(os.path.basename(src))[0]
    ext = os.path.splitext(src)[1] or '.xlsx'
    folder = outdir or os.path.dirname(os.path.abspath(src))
    cand = os.path.join(folder, base + SUFFIX + ext)
    i = 2
    while os.path.exists(cand):
        cand = os.path.join(folder, base + SUFFIX + "(%d)" % i + ext)
        i += 1
    return cand


# ---------------------------------------------------------------- 图形界面

def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    root = tk.Tk()
    root.title("%s v%s" % (APP_NAME, APP_VER))
    root.geometry("760x560")
    root.minsize(680, 500)

    # 选一个带中文的界面字体（Windows 用微软雅黑，其余平台依次回退）
    try:
        import tkinter.font as tkfont
        avail = set(tkfont.families())
        for fam in ("Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑",
                    "PingFang SC", "Noto Sans CJK SC", "WenQuanYi Zen Hei",
                    "SimHei", "宋体", "DejaVu Sans"):
            if fam in avail:
                for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont",
                             "TkHeadingFont", "TkCaptionFont", "TkTooltipFont"):
                    try:
                        tkfont.nametofont(name).configure(family=fam, size=10)
                    except Exception:
                        pass
                break
    except Exception:
        pass

    files = []
    mode = tk.StringVar(value='same')
    outdir = tk.StringVar(value='')
    outfile = tk.StringVar(value='')
    keep_links = tk.BooleanVar(value=True)
    running = {'v': False}

    pad = {'padx': 8, 'pady': 4}

    # 文件列表
    f1 = ttk.LabelFrame(root, text="① 要处理的 Excel 文件（原始文件只读，不会被修改）")
    f1.pack(fill='both', expand=False, **pad)
    lst = tk.Listbox(f1, height=6, selectmode='extended')
    lst.pack(side='left', fill='both', expand=True, padx=6, pady=6)
    sb = ttk.Scrollbar(f1, orient='vertical', command=lst.yview)
    sb.pack(side='left', fill='y', pady=6)
    lst.configure(yscrollcommand=sb.set)

    def refresh():
        lst.delete(0, 'end')
        for f in files:
            lst.insert('end', f)
        cnt.set("已选 %d 个文件" % len(files))

    def add_files():
        paths = filedialog.askopenfilenames(
            title="选择 Excel 文件",
            filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        for p in paths:
            if p not in files:
                files.append(p)
        refresh()

    def del_sel():
        for i in reversed(lst.curselection()):
            files.pop(i)
        refresh()

    def clear_all():
        files.clear()
        refresh()

    btns = ttk.Frame(f1)
    btns.pack(side='right', fill='y', padx=6, pady=6)
    ttk.Button(btns, text="添加文件…", command=add_files).pack(fill='x', pady=2)
    ttk.Button(btns, text="移除选中", command=del_sel).pack(fill='x', pady=2)
    ttk.Button(btns, text="清空列表", command=clear_all).pack(fill='x', pady=2)
    cnt = tk.StringVar(value="已选 0 个文件")
    ttk.Label(btns, textvariable=cnt).pack(fill='x', pady=(6, 0))

    # 输出设置
    f2 = ttk.LabelFrame(root, text="② 输出位置")
    f2.pack(fill='x', **pad)
    ttk.Radiobutton(f2, text="与原文件同目录（文件名自动加 “%s”）" % SUFFIX,
                    variable=mode, value='same').grid(row=0, column=0, sticky='w', padx=8, pady=3, columnspan=3)
    ttk.Radiobutton(f2, text="输出到指定文件夹：", variable=mode, value='dir'
                    ).grid(row=1, column=0, sticky='w', padx=8, pady=3)
    e_dir = ttk.Entry(f2, textvariable=outdir, width=52)
    e_dir.grid(row=1, column=1, sticky='we', pady=3)

    def pick_dir():
        d = filedialog.askdirectory(title="选择输出文件夹")
        if d:
            outdir.set(d)
            mode.set('dir')
    ttk.Button(f2, text="浏览…", command=pick_dir).grid(row=1, column=2, padx=6)

    ttk.Radiobutton(f2, text="保存为指定文件（仅单个源文件时）：", variable=mode, value='file'
                    ).grid(row=2, column=0, sticky='w', padx=8, pady=3)
    e_file = ttk.Entry(f2, textvariable=outfile, width=52)
    e_file.grid(row=2, column=1, sticky='we', pady=3)

    def pick_file():
        p = filedialog.asksaveasfilename(title="另存为", defaultextension=".xlsx",
                                        filetypes=[("Excel 工作簿", "*.xlsx")])
        if p:
            outfile.set(p)
            mode.set('file')
    ttk.Button(f2, text="浏览…", command=pick_file).grid(row=2, column=2, padx=6)
    f2.columnconfigure(1, weight=1)

    opt = ttk.Frame(root)
    opt.pack(fill='x', padx=8)
    ttk.Checkbutton(opt, text="保留对外部工作簿的引用（一般保留即可）",
                    variable=keep_links).pack(side='left')

    # 运行区
    f3 = ttk.Frame(root)
    f3.pack(fill='both', expand=True, **pad)
    bar = ttk.Progressbar(f3, mode='determinate', maximum=100)
    bar.pack(fill='x', pady=(2, 4))
    log = tk.Text(f3, height=12, wrap='word', state='disabled')
    log.pack(fill='both', expand=True)
    lgsb = ttk.Scrollbar(f3, orient='vertical', command=log.yview)
    lgsb.place(relx=1.0, rely=0, relheight=1, anchor='ne')
    log.configure(yscrollcommand=lgsb.set)

    def logln(s):
        log.configure(state='normal')
        log.insert('end', s + "\n")
        log.see('end')
        log.configure(state='disabled')
        root.update_idletasks()

    def set_prog(frac):
        bar['value'] = max(0, min(1, frac)) * 100

    def plan_outputs():
        """返回 [(src, dst)]，并做合法性检查。"""
        if not files:
            raise ValueError("请先添加要处理的 Excel 文件")
        m = mode.get()
        if m == 'file':
            if len(files) > 1:
                raise ValueError("“保存为指定文件”只适用于单个源文件；多个文件请选择文件夹")
            if not outfile.get().strip():
                raise ValueError("请填写输出文件名")
            if os.path.abspath(outfile.get()) == os.path.abspath(files[0]):
                raise ValueError("输出文件不能与原文件相同")
            return [(files[0], outfile.get().strip())]
        if m == 'dir':
            d = outdir.get().strip()
            if not d:
                raise ValueError("请填写或选择输出文件夹")
            os.makedirs(d, exist_ok=True)
        else:
            d = None
        return [(f, default_out(f, d)) for f in files]

    def worker(pairs):
        ok = 0
        try:
            for i, (src, dst) in enumerate(pairs):
                logln("=" * 60)
                logln("[%d/%d] %s" % (i + 1, len(pairs), os.path.basename(src)))
                base = i / len(pairs)
                span = 1 / len(pairs)
                pre = scan_workbook(src)
                logln("  工作表 %d 个，公式单元格 %d 个；内嵌图片单元格 %d 个（公式与格式原样保留）"
                      % (pre['sheets'], pre['formulas'], pre['keep_image']))
                if pre['no_cache']:
                    logln("  ⚠ 其中 %d 个公式在原文件里没有缓存计算结果（原文件从未被 Excel/WPS "
                          "保存过，或外部引用未取到数），这些格处理后会是空白。" % pre['no_cache'])
                st = convert(src, dst,
                             progress=lambda m, f=None: (
                                 set_prog(base + span * (f or 0)), logln("  " + m)),
                             keep_external=keep_links.get())
                logln("  ✔ 已剥离公式 %d 个；文本单元格 %d 个；保留图片单元格 %d 个（其公式与格式未动）；耗时 %.1fs"
                      % (st['formulas'], st['text'], st['keep_image'], st['seconds']))
                logln("  输出 → %s" % dst)
                ok += 1
            set_prog(1.0)
            logln("=" * 60)
            logln("全部完成：成功 %d / %d。原始文件均未改动。" % (ok, len(pairs)))
            messagebox.showinfo(APP_NAME, "处理完成：成功 %d / %d 个文件。\n\n原始文件未作任何修改。"
                                % (ok, len(pairs)))
        except Exception as e:
            logln("✘ 出错：%s" % e)
            logln(traceback.format_exc())
            messagebox.showerror(APP_NAME, "处理失败：\n%s" % e)
        finally:
            running['v'] = False
            btn_run.configure(state='normal')
            btn_open.configure(state='normal')

    def start():
        if running['v']:
            return
        try:
            pairs = plan_outputs()
        except ValueError as e:
            messagebox.showwarning(APP_NAME, str(e))
            return
        running['v'] = True
        btn_run.configure(state='disabled')
        set_prog(0)
        logln("开始处理 %d 个文件…" % len(pairs))
        threading.Thread(target=worker, args=(pairs,), daemon=True).start()

    def open_out():
        try:
            pairs = plan_outputs()
        except ValueError:
            pairs = [(f, default_out(f, None)) for f in files]
        d = os.path.dirname(pairs[0][1]) if pairs else ''
        if not d or not os.path.isdir(d):
            messagebox.showinfo(APP_NAME, "输出文件夹还不存在")
            return
        if sys.platform.startswith('win'):
            os.startfile(d)          # noqa: 仅 Windows
        elif sys.platform == 'darwin':
            os.system('open "%s"' % d)
        else:
            os.system('xdg-open "%s"' % d)

    bar2 = ttk.Frame(root)
    bar2.pack(fill='x', padx=8, pady=(0, 8))
    btn_run = ttk.Button(bar2, text="开始转换", command=start)
    btn_run.pack(side='left')
    btn_open = ttk.Button(bar2, text="打开输出文件夹", command=open_out)
    btn_open.pack(side='left', padx=6)
    ttk.Button(bar2, text="退出", command=root.destroy).pack(side='right')

    logln("%s v%s 就绪。原始文件全程只读，结果另存为新文件。" % (APP_NAME, APP_VER))
    logln("提示：本工具只搬运公式已有的计算结果，不负责计算。若图表需要保留，请勿用 Excel 另存再处理。")
    root.mainloop()
    return 0


def main():
    # PyInstaller --windowed 冻结环境下 stdout/stderr 可能是 None，
    # 此时 print/write 会抛异常，统一重定向到 devnull。
    for name in ('stdout', 'stderr'):
        if getattr(sys, name, None) is None:
            try:
                setattr(sys, name, open(os.devnull, 'w', encoding='utf-8'))
            except Exception:
                pass
    args = [a for a in sys.argv[1:]]
    if args and args[0] in ('--cli', 'cli'):
        return run_cli(args[1:])
    if args and args[0] in ('--help', '-h'):
        print(__doc__)
        return 0
    try:
        return run_gui()
    except Exception as e:                      # 没有 tkinter / 无图形环境时退回命令行
        sys.stderr.write("无法启动图形界面：%s\n" % e)
        sys.stderr.write("可用命令行模式：python %s --cli 源文件.xlsx [输出.xlsx]\n"
                         % os.path.basename(__file__))
        return 2


if __name__ == '__main__':
    sys.exit(main())
