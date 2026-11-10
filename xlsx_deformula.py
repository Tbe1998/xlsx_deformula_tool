#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 公式转数值工具 (Excel Formula to Value)

功能
----
1) 图形界面导入要处理的 Excel 文件：支持**拖拽**（可一次拖多个文件，也可整文件夹拖入），
   也保留「添加文件…」按钮（.xlsx / .xlsm）
2) 指定输出位置（源文件同目录自动命名 / 指定目录 / 指定单个文件名）
   · 输出文件夹**每次打开程序都要重新选择**，启动时不预填任何路径
   · 选过的输出文件夹会被记住（最多 10 个，最近优先），下次可从下拉历史直接选
3) 原始文件全程只读，绝不修改（只读取源文件，结果写入新文件）
4) 保留图表、图片、合并单元格、条件格式、行高列宽——只在 XML 层剥离 <f> 公式，
   保留公式的缓存计算结果 <v>，不用 openpyxl 重新保存（会丢图表）
5) 自带更新：启动 1.5 秒后后台静默检查 NAS 镜像上的 latest.json（局域网优先、
   公网 HTTPS 兜底），底栏「检查更新」可手动检查并一键下载安装（校验 sha256）；
   下载成功后自动关闭→替换→重启，全程隐藏执行（不弹 cmd 窗口），失败会弹提示并写日志

零第三方依赖：仅用 Python 标准库（tkinter + zipfile + re）。
可作为命令行工具使用：  python xlsx_deformula.py --cli 源文件.xlsx [输出.xlsx]
检查/安装更新：          python xlsx_deformula.py --check-update | --update | --version

注意：本工具不计算公式。若原文件从未被 Excel/WPS 保存过（公式没有缓存结果），
这些单元格处理后会变成空值——程序会在处理前统计并提示。
"""

import os
import re
import sys
import json
import time
import hashlib
import shutil
import zipfile
import tempfile
import threading
import subprocess
import traceback
import urllib.error
import urllib.parse
import urllib.request
from xml.sax.saxutils import escape

APP_NAME = "Excel 公式转数值工具"
APP_VER = "1.5"
APP_ID = "xlsx-deformula"
SUFFIX = "_无公式"

# 更新源（按序尝试）：家里的 NAS 镜像最快，公网 HTTPS 兜底
DEFAULT_MANIFESTS = [
    "http://192.168.1.6/download/%s/latest.json" % APP_ID,
    "https://www.topbei.cn/download/%s/latest.json" % APP_ID,
]
UA = "XlsxDeformula/%s" % APP_VER

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
    # 快路径：绝大多数公式块不含 DISPIMG，没必要为每个块都做一次反转义（大表里
    # 有几十万个公式块，这一步是实打实的 CPU 开销）。
    if 'DISPIMG' in block.upper() and DISPIMG_RE.search(_unescape(block)):
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
                    if 'DISPIMG' in blk.upper() and DISPIMG_RE.search(_unescape(blk)):
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
    out_dir = os.path.dirname(dst)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    tmp = dst + '.writing'
    try:
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
                    s = re.sub(r'\s*fullCalcOnLoad\s*=\s*"[^"]*"', '', s)
                    if not keep_external:
                        s = re.sub(r'<externalReferences>.*?</externalReferences>', '', s, flags=re.S)
                    data = s.encode('utf-8')
                zout.writestr(item, data)
                if progress and i % 20 == 0:
                    say("正在写入新文件… (%d/%d)" % (i, total), 0.35 + 0.5 * i / max(total, 1))
        shutil.move(tmp, dst)
    except Exception:
        # 写一半失败不能留下 .writing 垃圾文件（旧版本会残留，下次还覆盖同名临时文件）
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass
        raise

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
    try:
        pre = scan_workbook(src)
        print("源文件 : %s" % src)
        print("工作表 : %d   公式单元格: %d   其中无缓存值: %d   内嵌图片单元格(保留不动): %d"
              % (pre['sheets'], pre['formulas'], pre['no_cache'], pre['keep_image']))
        st = convert(src, dst, progress=lambda m, f=None: print("  " + m))
    except Exception as e:                                        # noqa: BLE001
        # 命令行模式也别甩 traceback 给用户；退出码 2 表示业务失败
        print("处理失败：%s" % e, file=sys.stderr)
        return 2
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


# ---------------------------------------------------------------- 配置 / 记忆
#
# 需求约定：
#   · 输出文件夹「每次打开程序都要重新选择」——启动时不预填任何路径，
#     指向"指定文件夹"的模式下若没选，转换会被拦下并提示；
#   · 历史输出文件夹会被记住（最多 10 条，最近优先、去重），
#     下拉即可直接选用，省去逐层点目录。

EXCEL_EXTS = ('.xlsx', '.xlsm')
HISTORY_MAX = 10


def config_file():
    """配置文件路径：Windows 用 %APPDATA%，其他平台用 ~/.config。"""
    if sys.platform.startswith('win'):
        base = os.environ.get('APPDATA') or os.path.expanduser('~')
    else:
        base = os.environ.get('XDG_CONFIG_HOME') or os.path.expanduser('~/.config')
    folder = os.path.join(base, 'XlsxDeformula')
    try:
        os.makedirs(folder, exist_ok=True)
    except Exception:
        folder = os.path.expanduser('~')
    return os.path.join(folder, 'config.json')


def load_config():
    import json
    path = config_file()
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return {'output_history': []}
        hist = [h for h in data.get('output_history', []) if isinstance(h, str) and h]
        data['output_history'] = hist
        return data
    except Exception:
        return {'output_history': []}


def save_config(cfg):
    import json
    try:
        with open(config_file(), 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def remember_output_dir(folder):
    """把文件夹记进历史（最近的排最前，最多 10 条，忽略大小写去重）。"""
    if not folder:
        return load_config()
    folder = os.path.abspath(folder)
    cfg = load_config()
    hist = [h for h in cfg.get('output_history', [])
            if os.path.normcase(os.path.abspath(h)) != os.path.normcase(folder)]
    hist.insert(0, folder)
    cfg['output_history'] = hist[:HISTORY_MAX]
    save_config(cfg)
    return cfg


def is_excel_file(path):
    name = os.path.basename(path)
    return (name.lower().endswith(EXCEL_EXTS)
            and not name.startswith('~$')          # Excel 临时文件
            and not name.startswith('.'))


def parse_dropped_paths(data, splitlist=None):
    """解析拖拽事件里的路径串。

    支持：单个文件、多个文件（含带空格的路径，Tk 会加花括号）、整个文件夹（收集其中的表格）。
    返回 (可处理的文件列表, 被忽略的条目列表)。
    """
    if splitlist is not None:
        try:
            parts = list(splitlist(data))
        except Exception:
            parts = [data]
    elif isinstance(data, str):
        # 没有 Tk splitlist 时的兜底：按花括号或空白切分，保留带空格的路径
        parts = re.findall(r'\{[^}]*\}|[^\s]+', data)
    else:
        parts = list(data)

    found, ignored = [], []
    for raw in parts:
        p = str(raw).strip()
        if len(p) >= 2 and p[0] == '{' and p[-1] == '}':
            p = p[1:-1]
        if len(p) >= 2 and p[0] == '"' and p[-1] == '"':
            p = p[1:-1]
        if not p:
            continue
        if os.path.isdir(p):
            hit = 0
            try:
                for name in sorted(os.listdir(p)):
                    full = os.path.join(p, name)
                    if os.path.isfile(full) and is_excel_file(full):
                        found.append(full)
                        hit += 1
            except Exception:
                pass
            if not hit:
                ignored.append(p)
        elif os.path.isfile(p) and is_excel_file(p):
            found.append(p)
        else:
            ignored.append(p)
    return found, ignored


# ---------------------------------------------------------------- 更新引擎
#
# 设计口径（与 SignStudio 一致）：
#   · 启动 1.5 秒后在后台静默检查，成功且无新版就只写一行日志（不弹窗、不打扰）
#   · 手动点「检查更新」才会把失败原因说清楚（内网/代理下连不上很正常）
#   · 下载边收边算 sha256，与清单不符直接拒绝安装
#   · 源码形态：解压覆盖程序目录（旧文件先备份到 .update_backup/）
#   · exe 形态：正在运行的 exe 在 Windows 上锁死，必须由独立 cmd 进程
#     等本进程退出后再替换文件、重新拉起

BACKUP_ROOT = os.path.join(os.path.dirname(os.path.abspath(sys.executable
                                                            if getattr(sys, 'frozen', False)
                                                            else __file__)), '.update_backup')


class UpdateError(RuntimeError):
    pass


def is_frozen():
    return bool(getattr(sys, 'frozen', False))


def program_dir():
    if is_frozen():
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _ascii_url(url):
    """URL 里的中文/空格必须百分号编码，否则 urllib 抛 ascii codec 错误。"""
    try:
        p = urllib.parse.urlsplit(url)
        path = urllib.parse.quote(p.path, safe="/%:@&=+$,-_.!~*'()")
        query = urllib.parse.quote(p.query, safe="=&%:/?@+")
        return urllib.parse.urlunsplit((p.scheme, p.netloc, path, query, p.fragment))
    except Exception:
        return url


def http_json(url, timeout=8):
    req = urllib.request.Request(_ascii_url(url), headers={
        "User-Agent": UA, "Accept": "application/json, */*",
        # 不要 gzip：urllib 不会自动解压，NAS/反代一开压缩就解析失败
        "Accept-Encoding": "identity", "Cache-Control": "no-cache"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            raw = r.read()
    except urllib.error.HTTPError as e:
        raise UpdateError("HTTP %s" % e.code)
    except urllib.error.URLError as e:
        raise UpdateError("连不上（%s）" % (getattr(e, "reason", e),))
    except Exception as e:                                     # noqa: BLE001
        raise UpdateError(str(e))
    try:
        return json.loads(raw.decode("utf-8", "replace"))
    except Exception:
        raise UpdateError("返回内容不是 JSON（可能被代理/强制门户拦了）")


def parse_version(text):
    t = str(text or "").strip().lstrip("vV")
    parts = []
    for seg in t.replace("-", ".").split("."):
        num = ""
        for ch in seg:
            if ch.isdigit():
                num += ch
            else:
                break
        parts.append(int(num) if num else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])


def is_newer(latest, current=APP_VER):
    return parse_version(latest) > parse_version(current)


def manifest_urls(custom=None):
    """更新源顺序：显式参数 > 环境变量 > 内置默认。"""
    for raw in (custom, os.environ.get("XLSX_DEFORMULA_MANIFESTS")):
        if raw:
            urls = [u.strip() for u in str(raw).replace(";", ",").split(",") if u.strip()]
            if urls:
                return urls
    return list(DEFAULT_MANIFESTS)


def check_latest(custom=None, timeout=8):
    """按顺序试各更新源，返回 {current, latest, has_update, notes, ...}；
    全部源都失败时抛 UpdateError。"""
    errors = []
    for url in manifest_urls(custom):
        host = urllib.parse.urlparse(url).netloc or url
        try:
            data = http_json(url, timeout=timeout)
        except Exception as e:                                  # noqa: BLE001
            errors.append("%s → %s" % (host, e))
            continue
        ver = str(data.get("version") or "").lstrip("vV")
        if not ver:
            errors.append("%s → 清单缺 version 字段" % host)
            continue
        info = {
            "current": APP_VER,
            "latest": ver,
            "has_update": is_newer(ver),
            "published": data.get("published") or "",
            "notes": (data.get("notes") or "").strip(),
            "source": url,
            "source_host": host,
            "checked_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        }
        for key in ("package", "exe"):
            name = data.get(key) or ""
            if not name:
                continue
            info[key] = name
            info[key + "_url"] = _ascii_url(urllib.parse.urljoin(url, name))
            info[key + "_size"] = int(data.get(key + "_size") or 0)
            info[key + "_sha256"] = (data.get(key + "_sha256") or "").strip().lower()
        if info["has_update"] and not (info.get("package_sha256") or info.get("exe_sha256")):
            info["has_update"] = False
            info["warning"] = "清单里没有 sha256，为安全起见不提示更新"
        return info
    raise UpdateError("所有更新源都连不上：\n" + "\n".join("  · " + x for x in errors))


def download_file(url, dest, progress=None, timeout=60, size_hint=0):
    """下载并实时算 sha256，返回 (字节数, sha256)。"""
    os.makedirs(os.path.dirname(dest) or ".", exist_ok=True)
    tmp = dest + ".part"
    h = hashlib.sha256()
    got = 0
    req = urllib.request.Request(_ascii_url(url), headers={"User-Agent": UA, "Accept": "*/*"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            total = int(r.headers.get("Content-Length") or 0)
            with open(tmp, "wb") as f:
                while True:
                    chunk = r.read(1 << 16)
                    if not chunk:
                        break
                    f.write(chunk)
                    h.update(chunk)
                    got += len(chunk)
                    if progress:
                        progress(got, total)
    except Exception as e:                                       # noqa: BLE001
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass
        raise UpdateError("下载失败：%s" % e)
    if got < 128:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise UpdateError("下载文件过小（%d 字节），可能被网络拦截" % got)
    os.replace(tmp, dest)
    # 清单给了字节数就核对：截断/被注入 HTML 的下载即使没有 sha256 也能挡住
    if size_hint and got != int(size_hint):
        try:
            os.remove(dest)
        except OSError:
            pass
        raise UpdateError("下载大小与清单不符（清单 %d 字节，实际 %d 字节），已拒绝安装"
                          % (int(size_hint), got))
    return got, h.hexdigest()


def _fetch(url, size_hint, sha_want, suffix, progress=None):
    dest = os.path.join(tempfile.gettempdir(),
                        "xlsx-deformula-update-%d%s" % (os.getpid(), suffix))
    size, sha = download_file(url, dest, progress=progress, size_hint=size_hint)
    if sha_want and sha != sha_want:
        try:
            os.remove(dest)
        except OSError:
            pass
        raise UpdateError("安装包 sha256 校验不一致（期望 %s…，实际 %s…），已拒绝安装"
                          % (sha_want[:12], sha[:12]))
    return dest, size, sha


def latest_payload(info):
    """挑出本形态该下载的东西：exe 形态优先用单文件 exe，源码形态用整包 zip。"""
    if is_frozen() and info.get("exe_url"):
        return info["exe_url"], info.get("exe_size") or 0, info.get("exe_sha256") or "", ".exe"
    if info.get("package_url"):
        return (info["package_url"], info.get("package_size") or 0,
                info.get("package_sha256") or "", ".zip")
    raise UpdateError("清单里没有可用于当前形态的安装包（缺 package / exe 字段）")


def apply_source_update(zip_path, version_tag=""):
    """源码形态：解压覆盖程序目录；旧文件备份到 .update_backup/<版本_时间>/。"""
    target = program_dir()
    stamp = "%s_%s" % (version_tag or "new", time.strftime("%Y%m%d_%H%M%S"))
    backup = os.path.join(BACKUP_ROOT, stamp)
    keep = {"config.json", ".update_backup", "__pycache__", ".buildenv", "build", "dist"}
    written = backed = 0
    try:
        with zipfile.ZipFile(zip_path) as zf:
            names = [m.filename.replace("\\", "/") for m in zf.infolist() if not m.is_dir()]
            for m in zf.infolist():
                name = m.filename.replace("\\", "/")
                if m.is_dir():
                    continue
                if name.startswith("/") or ".." in name.split("/"):
                    raise UpdateError("安装包里有非法路径：%s" % name)
                # Windows 盘符（C:/…）会被 os.path.join 当成绝对路径，写到程序目录之外
                if re.match(r"^[A-Za-z]:", name) or name.startswith("\\\\"):
                    raise UpdateError("安装包里有非法路径：%s" % name)
            tops = {n.split("/")[0] for n in names}
            prefix = (tops.pop() + "/") if len(tops) == 1 and any("/" in n for n in names) else ""
            for m in zf.infolist():
                name = m.filename.replace("\\", "/")
                if m.is_dir():
                    continue
                rel = name[len(prefix):] if prefix and name.startswith(prefix) else name
                rel = rel.lstrip("/")
                if not rel or rel.split("/")[0] in keep:
                    continue
                dst = os.path.join(target, rel.replace("/", os.sep))
                if os.path.isfile(dst):
                    bdst = os.path.join(backup, rel.replace("/", os.sep))
                    try:
                        os.makedirs(os.path.dirname(bdst), exist_ok=True)
                        shutil.copy2(dst, bdst)
                        backed += 1
                    except Exception:
                        pass
                os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
                with zf.open(m) as src, open(dst, "wb") as out:
                    shutil.copyfileobj(src, out)
                written += 1
    except UpdateError:
        raise
    except Exception as e:                                       # noqa: BLE001
        raise UpdateError("解压安装失败：%s" % e)
    return {"files": written, "backup": backup if backed else ""}


def _short_path(path):
    """Windows 8.3 短路径：让替换脚本保持纯 ASCII，绕开中文路径乱码。"""
    try:
        import ctypes
        buf = ctypes.create_unicode_buffer(1024)
        if ctypes.windll.kernel32.GetShortPathNameW(path, buf, 1024):
            return buf.value or path
    except Exception:
        pass
    return path


def apply_exe_update(new_exe, version_tag=""):
    """exe 形态：写一个隐藏的 VBS 替换脚本，等本进程退出后换掉 exe 并重新拉起。

    为什么不再用 cmd/bat：
      · cmd 窗口会闪一下，用户看到控制台却没后续，体验像是卡死；
      · bat 里 wait 父进程、move 失败都没有反馈，失败就静默结束。
    现在改成 wscript 隐藏执行 VBS：没有控制台窗口，替换失败会弹原生对话框并写日志。
    """
    if os.name != "nt":
        raise UpdateError("当前不是 Windows 可执行文件形态，无法自替换")
    target = os.path.abspath(sys.executable)
    tgt = _short_path(target)
    src = _short_path(new_exe)
    stem = os.path.join(tempfile.gettempdir(), "xlsx-deformula-update-%d" % os.getpid())
    vbs = stem + ".vbs"
    log = stem + ".log"
    text = _vbs_update_script(src, tgt, log, version_tag)
    with open(vbs, "w", encoding="utf-16", newline="\r\n") as f:
        f.write(text)

    flags = (getattr(subprocess, "CREATE_NO_WINDOW", 0)
             | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0))
    subprocess.Popen(["wscript.exe", "//nologo", vbs], close_fds=True,
                     creationflags=flags,
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return log


def _vbs_update_script(src, tgt, log, version_tag=""):
    """生成隐藏执行的 VBS 替换脚本（单独成函数便于自检）。"""
    def q(p):
        return '"%s"' % p.replace('"', '""')

    lines = [
        'Option Explicit',
        'Dim fso, sh, i, ok, msg',
        'Set fso = CreateObject("Scripting.FileSystemObject")',
        'Set sh  = CreateObject("WScript.Shell")',
        'Const SRC = %s' % q(src),
        'Const TGT = %s' % q(tgt),
        'Const LOG = %s' % q(log),
        'Const VER = %s' % q(str(version_tag or "")),
        '',
        'Sub Say(s)',
        '  On Error Resume Next',
        '  Dim f: Set f = fso.OpenTextFile(LOG, 8, True)',
        '  f.WriteLine Now & "  " & s',
        '  f.Close',
        'End Sub',
        '',
        'Say "update start, ver=" & VER',
        'ok = False',
        "' 直接覆盖复制：程序运行时目标 exe 被锁，CopyFile 会失败，正好当锁检测；",
        "' 进程退出后第一轮就成功，不需要先删旧文件（避免替换失败把程序弄丢）。",
        'For i = 1 To 60',
        '  On Error Resume Next',
        '  fso.CopyFile SRC, TGT, True',
        '  If Err.Number = 0 Then',
        '    If fso.GetFile(TGT).Size = fso.GetFile(SRC).Size Then ok = True',
        '  End If',
        '  Err.Clear',
        '  If ok Then Exit For',
        '  WScript.Sleep 1000',
        'Next',
        '',
        'If ok Then',
        '  On Error Resume Next',
        '  fso.DeleteFile SRC, True',
        '  Say "replaced ok after " & i & "s"',
        '  sh.Run """" & TGT & """", 1, False',
        '  Say "relaunched"',
        'Else',
        '  Say "FAILED after " & i & " tries"',
        '  msg = "更新文件替换失败，可能原因：" & vbCrLf & _',
        '        "· 程序仍在占用该文件（请手动关闭后重试）" & vbCrLf & _',
        '        "· 当前目录没有写入权限（试着把程序放到桌面或文档目录）" & vbCrLf & vbCrLf & _',
        '        "原程序会继续以旧版本启动。" & vbCrLf & _',
        '        "日志：" & LOG',
        '  MsgBox msg, 48, "Excel 公式转数值 - 更新未完成"',
        '  On Error Resume Next',
        '  sh.Run """" & TGT & """", 1, False',
        'End If',
        '',
        'On Error Resume Next',
        'fso.DeleteFile WScript.ScriptFullName, True',
    ]
    return "\r\n".join(lines) + "\r\n"


def install_update(info, progress=None):
    """下载 + 校验 + 安装。返回 (说明文字, 是否需要退出并重启)。"""
    url, size_hint, sha_want, suffix = latest_payload(info)
    path, got, sha = _fetch(url, size_hint, sha_want, suffix, progress=progress)
    try:
        if is_frozen() and suffix == ".exe":
            log = apply_exe_update(path, info.get("latest", ""))
            # exe 由 VBS 复制后自行删除；这里不删，避免替换过程中源文件先没了
            return ("新版本 v%s 已下载（%.1f MB）并校验通过。\n\n"
                    "程序将自动关闭 → 替换文件 → 重新启动，全程隐藏执行，不会弹出命令行窗口。\n"
                    "如 60 秒后仍未重启，更新日志在：\n%s"
                    % (info.get("latest"), got / 1048576.0, log)), True
        res = apply_source_update(path, info.get("latest", ""))
        return ("已覆盖 %d 个程序文件%s。程序将自动重启以载入 v%s。"
                % (res["files"], ("（旧版本备份在 %s）" % res["backup"]) if res["backup"] else "",
                   info.get("latest"))), True
    finally:
        # 源码形态安装完就把临时包删掉，别在 %TEMP% 里越堆越多
        if suffix == ".zip":
            try:
                os.remove(path)
            except OSError:
                pass



# ---------------------------------------------------------------- 图形界面
#
# 视觉规范（v1.3 起）：浅灰底 + 白色卡片 + 单一蓝色强调色，全部用 tkinter 原生
# 控件自绘，零第三方依赖，Windows / Linux 表现一致。
#   · 顶部：品牌标识 + 标题 + 版本徽标
#   · 主体：三张卡片（文件 / 输出位置 / 处理日志），编号圆标 + 充足留白
#   · 底部：白色操作条，主按钮蓝底白字，其余扁平描边
# 变量名沿用旧版（btn_run / btn_upd / lst / bar ...），--ui-check 自检依赖它们。

def run_gui(ui_check=False):
    import queue
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    # ---- 配色与字体 -------------------------------------------------
    BG = '#f4f6fb'          # 窗口底色
    CARD = '#ffffff'        # 卡片
    BORDER = '#e3e8f2'      # 分隔线 / 描边
    TEXT = '#1b2130'
    MUTED = '#707a8f'
    ACCENT = '#2f6bff'
    ACCENT_DARK = '#1f55d6'
    ACCENT_SOFT = '#eaf1ff'
    SUCCESS_C = '#0f8a5f'
    WARN_C = '#b26a00'
    ERROR_C = '#c93636'
    FIELD_BG = '#fbfcff'

    # 拖拽组件（tkdnd 的 Python 包装）；缺失或异常时不崩，退回按钮选文件。
    # 设置环境变量 XLSX_DEFORMULA_NO_DND=1 可强制关闭拖拽（排障用）。
    HAS_DND = False
    DND_FILES = None
    root = None
    if not os.environ.get('XLSX_DEFORMULA_NO_DND'):
        try:
            from tkinterdnd2 import TkinterDnD, DND_FILES as _DND_FILES
            root = TkinterDnD.Tk()
            DND_FILES = _DND_FILES
            HAS_DND = True
        except Exception as e:
            sys.stderr.write("拖拽组件不可用，改用按钮选文件：%s\n" % e)
            try:
                root.destroy()
            except Exception:
                pass
            root = None
    if root is None:
        root = tk.Tk()

    root.title("%s v%s" % (APP_NAME, APP_VER))
    # 初始尺寸只是占位，真正的高度在下面按控件实际需求自适应（见 fit_window()）
    root.geometry("900x800")
    root.minsize(780, 640)
    root.configure(bg=BG)

    # ---- 线程安全 UI 通道 -------------------------------------------
    # tkinter 只能在主线程操作控件。转换/更新跑在后台线程里，任何控件读写都
    # 必须投递到队列，由主线程 pump() 执行；否则 Windows 上会偶发卡死或
    # 报 "main thread is not in main loop"。root.after 在子线程里同样不可靠。
    ui_q = queue.Queue()

    def later(fn, *a):
        ui_q.put((fn, a))

    def on_main():
        return threading.current_thread() is threading.main_thread()

    def pump():
        while True:
            try:
                fn, a = ui_q.get_nowait()
            except queue.Empty:
                break
            try:
                fn(*a)
            except Exception:                                     # noqa: BLE001
                traceback.print_exc()
        root.after(120, pump)

    # 选一个带中文的界面字体（Windows 用微软雅黑，其余平台依次回退）
    FAM = 'TkDefaultFont'
    MONO = 'Courier New'
    try:
        import tkinter.font as tkfont
        avail = set(tkfont.families())
        for fam in ("Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑",
                    "PingFang SC", "Noto Sans SC", "Noto Sans CJK SC",
                    "WenQuanYi Zen Hei", "wenquanyi bitmap song",
                    "WenQuanYi Bitmap Song", "SimHei", "宋体", "DejaVu Sans"):
            if fam in avail:
                FAM = fam
                for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont",
                             "TkHeadingFont", "TkCaptionFont", "TkTooltipFont"):
                    try:
                        tkfont.nametofont(name).configure(family=fam, size=10)
                    except Exception:
                        pass
                break
        for fam in ("Consolas", "Cascadia Mono", "JetBrains Mono", "Menlo",
                    "DejaVu Sans Mono", "Noto Sans Mono", "Courier New"):
            if fam in avail:
                MONO = fam
                break
    except Exception:
        pass

    # ---- ttk 主题：统一成扁平浅色 ----------------------------------
    style = ttk.Style(root)
    try:
        style.theme_use('clam')
    except Exception:
        pass
    style.configure('.', background=BG, foreground=TEXT, font=(FAM, 10))
    style.configure('Card.TFrame', background=CARD)
    style.configure('Card.TLabel', background=CARD, foreground=TEXT)
    style.configure('Card.TRadiobutton', background=CARD, foreground=TEXT,
                    indicatorcolor=FIELD_BG, focuscolor=CARD, padding=(0, 2))
    style.map('Card.TRadiobutton',
              background=[('active', CARD)],
              indicatorcolor=[('selected', ACCENT), ('active', ACCENT_SOFT),
                              ('pressed', ACCENT)])
    style.configure('Card.TCheckbutton', background=CARD, foreground=TEXT,
                    indicatorcolor=FIELD_BG, focuscolor=CARD, padding=(0, 2))
    style.map('Card.TCheckbutton',
              background=[('active', CARD)],
              indicatorcolor=[('selected', ACCENT), ('active', ACCENT_SOFT)])
    style.configure('Field.TEntry', fieldbackground=FIELD_BG, foreground=TEXT,
                    bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
                    insertcolor=TEXT, padding=6)
    style.map('Field.TEntry', bordercolor=[('focus', ACCENT)],
              lightcolor=[('focus', ACCENT)], darkcolor=[('focus', ACCENT)])
    style.configure('Field.TCombobox', fieldbackground=FIELD_BG, background=FIELD_BG,
                    foreground=TEXT, bordercolor=BORDER, lightcolor=BORDER,
                    darkcolor=BORDER, arrowcolor=MUTED, padding=5)
    style.map('Field.TCombobox',
              fieldbackground=[('readonly', FIELD_BG), ('focus', '#ffffff')],
              bordercolor=[('focus', ACCENT)],
              lightcolor=[('focus', ACCENT)], darkcolor=[('focus', ACCENT)])
    style.configure('Accent.Horizontal.TProgressbar', troughcolor='#e9eef8',
                    background=ACCENT, bordercolor='#e9eef8',
                    lightcolor=ACCENT, darkcolor=ACCENT, thickness=6)
    style.configure('Card.Vertical.TScrollbar', background='#c9d2e4',
                    troughcolor=CARD, bordercolor=CARD, arrowcolor=MUTED,
                    lightcolor=CARD, darkcolor=CARD, width=10)
    root.option_add('*TCombobox*Listbox.background', CARD)
    root.option_add('*TCombobox*Listbox.foreground', TEXT)
    root.option_add('*TCombobox*Listbox.selectBackground', ACCENT_SOFT)
    root.option_add('*TCombobox*Listbox.selectForeground', TEXT)
    root.option_add('*TCombobox*Listbox.borderWidth', 0)

    # ---- 自绘控件 ---------------------------------------------------
    def round_rect(cv, x1, y1, x2, y2, r, **kw):
        pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
               x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
               x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
        return cv.create_polygon(pts, smooth=True, **kw)

    def flat_button(parent, text, command, kind='secondary', **kw):
        """扁平按钮：外层 1px 描边框 + 内层无边框按钮，保证边框颜色可控。"""
        fills = {
            'primary': (ACCENT, '#ffffff', ACCENT_DARK),
            'secondary': ('#ffffff', TEXT, ACCENT_SOFT),
            'ghost': (CARD, MUTED, '#eef2f9'),
        }
        fill, fg, hover = fills.get(kind, fills['secondary'])
        edge = ACCENT if kind == 'primary' else (BORDER if kind == 'secondary' else CARD)
        holder = tk.Frame(parent, bg=edge, bd=0, highlightthickness=0)
        btn = tk.Button(holder, text=text, command=command, relief='flat', bd=0,
                        bg=fill, fg=fg, activebackground=hover, activeforeground=fg,
                        disabledforeground='#9aa6bd', cursor='hand2',
                        padx=kw.get('padx', 16), pady=kw.get('pady', 6),
                        font=(FAM, 10, 'bold' if kind == 'primary' else 'normal'),
                        highlightthickness=0, takefocus=0)
        btn.pack(fill='both', expand=True, padx=1, pady=1)

        def on_enter(_e=None):
            if str(btn.cget('state')) != 'disabled':
                btn.configure(bg=hover)

        def on_leave(_e=None):
            btn.configure(bg=holder.fill)

        btn.bind('<Enter>', on_enter)
        btn.bind('<Leave>', on_leave)
        holder.btn = btn
        holder.fill = fill

        def configure(**k):
            direct = {kk: vv for kk, vv in k.items()
                      if kk in ('text', 'state', 'command', 'bg', 'fg', 'cursor')}
            if 'bg' in direct:
                holder.fill = direct['bg']
            rest = {kk: vv for kk, vv in k.items() if kk not in direct}
            if direct:
                btn.configure(**direct)
            if rest:
                tk.Frame.configure(holder, **rest)

        def cget(key):
            if key in ('text', 'state', 'bg', 'fg'):
                return btn.cget(key)
            return tk.Frame.cget(holder, key)

        def keys():
            ks = list(tk.Frame.keys(holder))
            return ks + [k for k in ('text', 'state') if k not in ks]

        holder.configure = configure
        holder.config = configure
        holder.cget = cget
        holder.keys = keys
        return holder

    def make_card(parent, num, title, hint=''):
        """带编号徽标的白色卡片。返回 (卡片, 内容区)。"""
        card = tk.Frame(parent, bg=CARD, highlightbackground=BORDER,
                        highlightcolor=BORDER, highlightthickness=1, bd=0)
        head = tk.Frame(card, bg=CARD)
        head.pack(fill='x', padx=14, pady=(10, 6))
        badge = tk.Canvas(head, width=20, height=20, bg=CARD, highlightthickness=0)
        badge.create_oval(0, 0, 19, 19, fill=ACCENT_SOFT, outline='')
        badge.create_text(10, 10, text=str(num), fill=ACCENT,
                          font=(FAM, 8, 'bold'))
        badge.pack(side='left')
        tk.Label(head, text=title, bg=CARD, fg=TEXT,
                 font=(FAM, 11, 'bold')).pack(side='left', padx=(8, 0))
        if hint:
            tk.Label(head, text=hint, bg=CARD, fg=MUTED,
                     font=(FAM, 9)).pack(side='right')
        body = tk.Frame(card, bg=CARD)
        body.pack(fill='both', expand=True, padx=14, pady=(0, 10))
        return card, body

    def choice_row(parent, kind, text, variable, value=True):
        """自绘单选/复选框：Canvas 指示器 + 文字，点击整行切换。

        不用 ttk 自带 indicator —— clam 主题下的勾选符依赖字体，缺字时会变成方块。
        """
        row = tk.Frame(parent, bg=CARD)
        cv = tk.Canvas(row, width=18, height=18, bg=CARD,
                       highlightthickness=0, cursor='hand2')
        cv.pack(side='left', padx=(0, 8))
        lbl = tk.Label(row, text=text, bg=CARD, fg=TEXT, font=(FAM, 10),
                       cursor='hand2')
        lbl.pack(side='left')

        def redraw(_a=None, _b=None, _c=None):
            on = (variable.get() == value) if kind == 'radio' else bool(variable.get())
            cv.delete('all')
            if kind == 'radio':
                cv.create_oval(1.5, 1.5, 16.5, 16.5, fill='#ffffff',
                               outline=ACCENT if on else '#b9c3d6', width=2)
                if on:
                    cv.create_oval(5.5, 5.5, 12.5, 12.5, fill=ACCENT, outline='')
            elif on:
                cv.create_rectangle(1, 1, 17, 17, fill=ACCENT, outline=ACCENT, width=1)
                cv.create_line(5, 9, 8, 12.5, 13, 5.5, fill='#ffffff', width=2,
                               capstyle='round', joinstyle='round')
            else:
                cv.create_rectangle(1, 1, 17, 17, fill='#ffffff',
                                    outline='#b9c3d6', width=1)
            return on

        def click(_e=None):
            if kind == 'radio':
                variable.set(value)
            else:
                variable.set(not bool(variable.get()))
            return 'break'

        def hover(on):
            if kind == 'radio':
                if variable.get() != value:
                    cv.itemconfigure(1, outline=ACCENT if on else '#b9c3d6')
            elif not variable.get():
                cv.itemconfigure(1, outline=ACCENT if on else '#b9c3d6')

        for w in (row, cv, lbl):
            w.bind('<Button-1>', click)
            w.bind('<Enter>', lambda e: hover(True))
            w.bind('<Leave>', lambda e: hover(False))
        try:
            variable.trace_add('write', redraw)
        except Exception:
            pass
        redraw()
        return row


    cfg = load_config()
    files = []
    mode = tk.StringVar(value='same')
    outdir = tk.StringVar(value='')          # 每次启动都从空开始：必须重新选择
    outfile = tk.StringVar(value='')
    keep_links = tk.BooleanVar(value=True)
    running = {'v': False}

    # ---------- 顶部品牌区 ----------
    header = tk.Frame(root, bg=CARD)
    header.pack(fill='x')
    hrow = tk.Frame(header, bg=CARD)
    hrow.pack(fill='x', padx=18, pady=(14, 12))

    logo = tk.Canvas(hrow, width=40, height=40, bg=CARD, highlightthickness=0)
    logo.pack(side='left')
    round_rect(logo, 0, 0, 39, 39, 10, fill=ACCENT, outline='')
    logo.create_text(20, 20, text='fx', fill='#ffffff', font=(FAM, 15, 'bold'))

    title_box = tk.Frame(hrow, bg=CARD)
    title_box.pack(side='left', padx=(12, 0))
    tk.Label(title_box, text=APP_NAME, bg=CARD, fg=TEXT,
             font=(FAM, 15, 'bold')).pack(anchor='w')
    subtitle = tk.Label(title_box,
                        text='公式 → 数值 · 保留图表 / 图片 / 格式 · 原始文件只读',
                        bg=CARD, fg=MUTED, font=(FAM, 9))
    subtitle.pack(anchor='w', pady=(2, 0))

    tk.Label(hrow, text='v%s' % APP_VER, bg=ACCENT_SOFT, fg=ACCENT,
             font=(FAM, 9, 'bold'), padx=10, pady=3).pack(side='right')
    tk.Frame(root, bg=BORDER, height=1).pack(fill='x')

    # ---------- 底部操作条（先占位，窗口再矮也不会被挤掉） ----------
    footer = tk.Frame(root, bg=CARD)
    footer.pack(side='bottom', fill='x')
    tk.Frame(footer, bg=BORDER, height=1).pack(fill='x')
    bar2 = tk.Frame(footer, bg=CARD)
    bar2.pack(fill='x', padx=14, pady=10)

    # ---------- 内容区 ----------
    content = tk.Frame(root, bg=BG)
    content.pack(fill='both', expand=True, padx=14, pady=12)

    # ---------- ① 文件 ----------
    card1, body1 = make_card(content, 1, '要处理的 Excel 文件',
                             '拖入或选择 · .xlsx / .xlsm')
    card1.pack(fill='x')

    hover_state = {'v': False}
    if HAS_DND:
        drop_title = '把 Excel 文件（或整个文件夹）拖到这里'
        drop_sub = '支持一次多个 · 也可以点击这里或「添加文件」按钮选择'
    else:
        drop_title = '当前运行环境缺少拖拽组件'
        drop_sub = '请点击这里或「添加文件」按钮选择 Excel 文件'

    drop_zone = tk.Canvas(body1, height=82, bg=CARD, highlightthickness=0,
                          cursor='hand2')
    drop_zone.pack(fill='x')

    def draw_drop(*_a):
        w = max(drop_zone.winfo_width(), 240)
        h = 82
        drop_zone.delete('all')
        hot = hover_state['v']
        drop_zone.create_rectangle(1, 1, w - 1, h - 1,
                                   fill='#eef4ff' if hot else '#fafbff',
                                   outline=ACCENT if hot else '#cbd8ef',
                                   dash=(6, 4), width=1)
        cx = w / 2
        cy = 23
        drop_zone.create_oval(cx - 15, cy - 15, cx + 15, cy + 15,
                              fill=ACCENT_SOFT, outline='')
        drop_zone.create_line(cx, cy - 7, cx, cy + 4, fill=ACCENT, width=2,
                              capstyle='round')
        drop_zone.create_line(cx - 5, cy - 1, cx, cy + 5, cx + 5, cy - 1,
                              fill=ACCENT, width=2, capstyle='round',
                              joinstyle='round')
        drop_zone.create_line(cx - 6, cy + 7, cx + 6, cy + 7, fill=ACCENT,
                              width=2, capstyle='round')
        drop_zone.create_text(cx, 54, text=drop_title, fill=TEXT,
                              font=(FAM, 10, 'bold'))
        drop_zone.create_text(cx, 70, text=drop_sub, fill=MUTED, font=(FAM, 9))

    drop_zone.bind('<Configure>', draw_drop)
    drop_zone.bind('<Enter>',
                   lambda e: (hover_state.__setitem__('v', True), draw_drop()))
    drop_zone.bind('<Leave>',
                   lambda e: (hover_state.__setitem__('v', False), draw_drop()))

    list_row = tk.Frame(body1, bg=CARD)
    list_row.pack(fill='both', expand=True, pady=(10, 0))

    list_wrap = tk.Frame(list_row, bg=FIELD_BG,
                         highlightbackground=BORDER, highlightthickness=1)
    list_wrap.pack(side='left', fill='both', expand=True)

    lst = tk.Listbox(list_wrap, height=5, bd=0, relief='flat',
                     highlightthickness=0, bg=FIELD_BG, fg=TEXT,
                     selectmode='extended', activestyle='none',
                     selectbackground=ACCENT_SOFT, selectforeground=TEXT,
                     font=(FAM, 9))
    lst.pack(side='left', fill='both', expand=True, padx=(8, 0), pady=6)
    sb = ttk.Scrollbar(list_wrap, orient='vertical', command=lst.yview,
                       style='Card.Vertical.TScrollbar')
    sb.pack(side='right', fill='y', padx=(2, 4), pady=6)
    lst.configure(yscrollcommand=sb.set)

    empty_lbl = tk.Label(list_wrap, text='还没有添加文件\n拖入或点「添加文件」开始',
                         bg=FIELD_BG, fg=MUTED, justify='center',
                         font=(FAM, 9))
    empty_lbl.place(relx=0.5, rely=0.5, anchor='center')
    empty_lbl.lift()

    cnt = tk.StringVar(value="已选 0 个文件")
    log_holder = {}

    def logln(s):
        # 后台线程（转换/更新）调用时投递回主线程，别直接碰 Text 控件
        if not on_main():
            later(logln, s)
            return
        box = log_holder.get('box')
        if box is None:
            return
        t = s.lstrip()
        tag = ''
        if t.startswith('='):
            tag = 'dim'
        elif t.startswith('✔'):
            tag = 'ok'
        elif t.startswith('✘'):
            tag = 'err'
        elif '⚠' in t:
            tag = 'warn'
        elif '输出 →' in t:
            tag = 'path'
        elif t.startswith('[') and ']' in t:
            tag = 'head'
        box.configure(state='normal')
        box.insert('end', s + "\n", tag)
        box.see('end')
        box.configure(state='disabled')
        root.update_idletasks()

    def refresh():
        lst.delete(0, 'end')
        for i, f in enumerate(files):
            lst.insert('end', f)
            if i % 2:
                lst.itemconfig(i, background='#f6f9fe')   # 斑马纹
        cnt.set("已选 %d 个文件" % len(files))
        if files:
            empty_lbl.place_forget()
        else:
            empty_lbl.place(relx=0.5, rely=0.5, anchor='center')
            empty_lbl.lift()

    def add_paths(paths, rejected=()):
        added = 0
        for p in paths:
            p = os.path.abspath(p)
            if p not in files:
                files.append(p)
                added += 1
        refresh()
        if added:
            logln("已加入 %d 个文件。" % added)
        if rejected:
            shown = "、".join(os.path.basename(r) for r in rejected[:5])
            more = "" if len(rejected) <= 5 else " 等 %d 项" % len(rejected)
            logln("已忽略非 Excel 条目：%s%s" % (shown, more))
            messagebox.showwarning(APP_NAME, "以下条目不是 .xlsx / .xlsm 表格，已忽略：\n\n%s%s"
                                   % (shown, more))

    def add_files():
        paths = filedialog.askopenfilenames(
            title="选择 Excel 文件",
            filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if paths:
            add_paths(paths)

    def del_sel():
        for i in reversed(lst.curselection()):
            files.pop(i)
        refresh()

    def clear_all():
        files.clear()
        refresh()

    def on_drop(event):
        found, ignored = parse_dropped_paths(event.data, root.tk.splitlist)
        if found or ignored:
            add_paths(found, ignored)

    btn_col = tk.Frame(list_row, bg=CARD)
    btn_col.pack(side='right', fill='y', padx=(10, 0))
    flat_button(btn_col, '添加文件', add_files, 'primary',
                pady=6).pack(fill='x')
    flat_button(btn_col, '移除选中', del_sel, pady=5).pack(fill='x', pady=(6, 0))
    flat_button(btn_col, '清空列表', clear_all, pady=5).pack(fill='x', pady=(6, 0))
    tk.Label(btn_col, textvariable=cnt, bg=CARD, fg=MUTED,
             font=(FAM, 9)).pack(anchor='w', pady=(10, 0))

    empty_lbl.bind('<Button-1>', lambda e: add_files())
    drop_zone.bind('<Button-1>', lambda e: add_files())

    if HAS_DND:
        for widget in (root, drop_zone, lst):
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind('<<Drop>>', on_drop)
            except Exception:
                pass

    # ---------- ② 输出位置 ----------
    card2, body2 = make_card(content, 2, '输出位置', '每次启动需重新选择')
    card2.pack(fill='x', pady=(12, 0))
    body2.columnconfigure(1, weight=1)

    choice_row(body2, 'radio', "与原文件同目录（文件名自动加「%s」）" % SUFFIX,
               mode, 'same').grid(row=0, column=0, columnspan=3, sticky='w',
                                  pady=(0, 2))
    choice_row(body2, 'radio', "输出到指定文件夹：", mode, 'dir').grid(
        row=1, column=0, sticky='w', padx=(0, 10), pady=3)

    hist = list(cfg.get('output_history', []))
    cb_dir = ttk.Combobox(body2, textvariable=outdir, values=hist, width=52,
                          style='Field.TCombobox')
    cb_dir.grid(row=1, column=1, sticky='we', pady=3)

    def on_hist(_evt=None):
        if outdir.get().strip():
            mode.set('dir')

    cb_dir.bind('<<ComboboxSelected>>', on_hist)

    def pick_dir():
        start = outdir.get().strip() or (hist[0] if hist else '')
        if not os.path.isdir(start):
            start = os.path.expanduser('~')
        d = filedialog.askdirectory(title="选择输出文件夹", initialdir=start)
        if d:
            outdir.set(d)
            mode.set('dir')

    flat_button(body2, '浏览…', pick_dir, padx=12, pady=3).grid(
        row=1, column=2, padx=(8, 0), pady=3)

    choice_row(body2, 'radio', "保存为指定文件（仅单个源文件时）：", mode, 'file').grid(
        row=2, column=0, sticky='w', padx=(0, 10), pady=3)
    ttk.Entry(body2, textvariable=outfile, width=52,
              style='Field.TEntry').grid(row=2, column=1, sticky='we', pady=3)

    def pick_file():
        p = filedialog.asksaveasfilename(title="另存为", defaultextension=".xlsx",
                                         filetypes=[("Excel 工作簿", "*.xlsx")])
        if p:
            outfile.set(p)
            mode.set('file')

    flat_button(body2, '浏览…', pick_file, padx=12, pady=3).grid(
        row=2, column=2, padx=(8, 0), pady=3)

    hist_state = tk.StringVar(
        value=("历史输出文件夹：%d 个，可下拉速选" % len(hist)) if hist
        else "历史输出文件夹：暂无（选过自动记住，最多 10 个）")

    hint_bar = tk.Frame(body2, bg='#fff8e8')
    hint_bar.grid(row=3, column=0, columnspan=3, sticky='we', pady=(6, 0))
    tk.Label(hint_bar, textvariable=hist_state, bg='#fff8e8', fg='#8a5a00',
             font=(FAM, 9)).pack(side='left', padx=10, pady=4)
    tk.Label(hint_bar, text='每次打开程序都需重新选择输出文件夹',
             bg='#fff8e8', fg='#b07a2a', font=(FAM, 9)).pack(
        side='right', padx=10, pady=4)

    choice_row(body2, 'check', "保留对外部工作簿的引用（一般保留即可）",
               keep_links).grid(row=5, column=0, columnspan=3, sticky='w',
                                pady=(6, 0))

    # ---------- ③ 运行区 ----------
    card3, body3 = make_card(content, 3, '处理日志', '原始文件全程只读')
    card3.pack(fill='both', expand=True, pady=(12, 0))

    bar = ttk.Progressbar(body3, mode='determinate', maximum=100,
                          style='Accent.Horizontal.TProgressbar')
    bar.pack(fill='x', pady=(0, 8))

    log_wrap = tk.Frame(body3, bg=CARD)
    log_wrap.pack(fill='both', expand=True)
    log_box = tk.Text(log_wrap, height=8, wrap='word', state='disabled', bd=0,
                      relief='flat', bg=FIELD_BG, fg='#39415a',
                      insertbackground=TEXT, font=(MONO, 9), padx=10, pady=8,
                      highlightthickness=1, highlightbackground=BORDER,
                      highlightcolor=BORDER, selectbackground=ACCENT_SOFT,
                      selectforeground=TEXT)
    log_box.pack(fill='both', expand=True, side='left')
    lgsb = ttk.Scrollbar(log_wrap, orient='vertical', command=log_box.yview,
                         style='Card.Vertical.TScrollbar')
    lgsb.pack(side='right', fill='y')
    log_box.configure(yscrollcommand=lgsb.set)
    log_box.tag_configure('ok', foreground=SUCCESS_C)
    log_box.tag_configure('err', foreground=ERROR_C)
    log_box.tag_configure('warn', foreground=WARN_C)
    log_box.tag_configure('path', foreground=ACCENT)
    log_box.tag_configure('dim', foreground='#c2c9d8')
    log_box.tag_configure('head', foreground=TEXT, font=(MONO, 9, 'bold'))
    log_holder['box'] = log_box

    def set_prog(frac):
        if not on_main():
            later(set_prog, frac)
            return
        bar['value'] = max(0, min(1, frac)) * 100

    def history_values():
        return list(load_config().get('output_history', []))

    def plan_outputs():
        if not files:
            raise ValueError("请先添加或拖入要处理的 Excel 文件")
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
                raise ValueError("本次启动还没有选择输出文件夹。\n"
                                 "请点「浏览…」选一个，或从下拉的历史记录里直接挑一个。")
            if not os.path.isdir(d):
                try:
                    os.makedirs(d, exist_ok=True)
                except Exception as e:
                    raise ValueError("输出文件夹不可用：%s\n%s" % (d, e))
            remember_output_dir(d)
            vals = history_values()
            cb_dir.configure(values=vals)
            hist_state.set("历史输出文件夹：%d 个，可下拉直接选" % len(vals))
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
            later(messagebox.showinfo, APP_NAME,
                  "处理完成：成功 %d / %d 个文件。\n\n原始文件未作任何修改。" % (ok, len(pairs)))
        except Exception as e:
            logln("✘ 出错：%s" % e)
            logln(traceback.format_exc())
            later(messagebox.showerror, APP_NAME, "处理失败：\n%s" % e)
        finally:
            running['v'] = False
            later(restore_buttons)

    def restore_buttons():
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

    btn_run = flat_button(bar2, '开始转换', start, 'primary', padx=22, pady=7)
    btn_run.pack(side='left')
    btn_open = flat_button(bar2, '打开输出文件夹', open_out)
    btn_open.pack(side='left', padx=(8, 0))

    # ---------- 更新（手动检查 + 启动静默检查） ----------
    upd = {'info': None, 'busy': False}

    def upd_status(text):
        ver_lbl.configure(text=text)

    def relaunch_self():
        """源码形态：更新完直接拉起新进程；exe 形态交给 VBS 替换脚本。"""
        if is_frozen():
            return
        try:
            subprocess.Popen([sys.executable, os.path.abspath(sys.argv[0])],
                             cwd=program_dir(), close_fds=True,
                             creationflags=getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0))
        except Exception:                                         # noqa: BLE001
            traceback.print_exc()

    def shutdown_and_exit():
        """先销毁窗口再硬退出，保证文件锁立刻释放，替换脚本不会卡在等待。"""
        try:
            root.quit()
            root.destroy()
        except Exception:                                         # noqa: BLE001
            pass
        os._exit(0)

    def run_update(info):
        """（后台线程）下载 → 校验 → 安装。"""
        def prog(got, total):
            if total:
                set_prog(got / float(total))
        try:
            msg, need_exit = install_update(info, progress=prog)
        except Exception as e:                                   # noqa: BLE001
            err = str(e)
            upd['busy'] = False
            later(upd_status, "v%s" % APP_VER)
            later(lambda: btn_upd.configure(state='normal', text="检查更新"))
            later(logln, "✘ 更新失败：%s" % err)
            later(lambda: messagebox.showerror(APP_NAME, "更新失败：\n%s" % err))
            return
        upd['busy'] = False
        later(set_prog, 1.0)
        later(logln, "✔ " + msg)
        if need_exit:
            later(lambda: messagebox.showinfo(APP_NAME, msg))
            later(relaunch_self)
            later(shutdown_and_exit)
        else:
            later(lambda: btn_upd.configure(state='normal', text="检查更新"))
            later(upd_status, "v%s" % APP_VER)
            later(lambda: messagebox.showinfo(APP_NAME, msg))

    def do_check(manual=False):
        if upd['busy']:
            return
        upd['busy'] = True
        if manual:
            btn_upd.configure(state='disabled', text="正在检查…")
            logln("正在检查更新…")

        def job():
            try:
                info = check_latest(timeout=8)
            except Exception as e:                               # noqa: BLE001
                err = str(e)
                upd['busy'] = False
                if manual:
                    later(lambda: btn_upd.configure(state='normal', text="检查更新"))
                    later(logln, "检查更新失败：%s" % err)
                    later(lambda: messagebox.showwarning(
                        APP_NAME, "检查更新失败（离线/内网被拦也算正常）：\n%s" % err))
                return
            if not info['has_update']:
                upd['busy'] = False
                later(logln, "已是最新版本 v%s。" % APP_VER)
                if manual:
                    later(lambda: btn_upd.configure(state='normal', text="检查更新"))
                    later(lambda: messagebox.showinfo(
                        APP_NAME, "当前已是最新版本 v%s。" % APP_VER))
                return
            upd['info'] = info
            later(logln, "发现新版本 v%s（当前 v%s）。" % (info['latest'], APP_VER))
            if info.get('notes'):
                later(logln, "更新说明：" + info['notes'].replace("\n", " / "))
            if not manual:
                upd['busy'] = False
                later(lambda: btn_upd.configure(text="更新到 v%s" % info['latest']))
                return
            later(lambda: confirm(info))

        threading.Thread(target=job, daemon=True).start()

    def confirm(info):
        ask = ("发现新版本 v%s（当前 v%s）\n\n%s\n\n现在下载并更新吗？"
               % (info['latest'], APP_VER, info.get('notes') or "（无更新说明）"))
        if not messagebox.askyesno(APP_NAME, ask):
            upd['busy'] = False
            btn_upd.configure(state='normal', text="检查更新")
            return
        btn_upd.configure(state='disabled', text="正在更新…")
        set_prog(0)
        threading.Thread(target=run_update, args=(info,), daemon=True).start()

    def on_update_click():
        if upd['info'] and upd['info'].get('has_update') and not upd['busy']:
            upd['busy'] = True
            confirm(upd['info'])
        else:
            do_check(manual=True)

    btn_exit = flat_button(bar2, '退出', root.destroy, 'ghost', padx=14, pady=6)
    btn_exit.pack(side='right')
    ver_lbl = tk.Label(bar2, text="v%s" % APP_VER, bg=CARD, fg=MUTED,
                       font=(FAM, 9))
    ver_lbl.pack(side='right', padx=(0, 12))
    btn_upd = flat_button(bar2, '检查更新', on_update_click, 'ghost',
                          padx=14, pady=6)
    btn_upd.pack(side='right', padx=(0, 14))

    def fit_window():
        """按控件实际需求高宽定窗口，保证所有按钮（含最底部一排）都在屏幕内。"""
        root.update_idletasks()
        want_w = max(root.winfo_reqwidth(), 780)
        want_h = root.winfo_reqheight()
        sw = root.winfo_screenwidth()
        shp = root.winfo_screenheight()
        w = min(want_w, max(sw - 60, 640))
        h_max = max(shp - 90, 560)                  # 留出任务栏/标题栏
        # 屏幕不够高时优先压缩日志框，而不是把底部按钮挤出屏幕
        for _ in range(10):
            root.update_idletasks()
            want_h = root.winfo_reqheight()
            if want_h <= h_max:
                break
            try:
                cur = int(log_box.cget('height'))
            except Exception:
                cur = 0
            if cur <= 5:
                break
            log_box.configure(height=max(5, cur - 1))
        # 仍然放不下时再压缩文件列表（日志已到最小可读高度）
        for _ in range(3):
            root.update_idletasks()
            if root.winfo_reqheight() <= h_max:
                break
            try:
                cur = int(lst.cget('height'))
            except Exception:
                cur = 0
            if cur <= 3:
                break
            lst.configure(height=cur - 1)
        # 小屏（如 1366×768）仍放不下时，收起拖拽区，保住日志与底部按钮
        if root.winfo_reqheight() > h_max:
            try:
                if drop_zone.winfo_ismapped():
                    drop_zone.pack_forget()
                    root.update_idletasks()
            except Exception:
                pass
        # 再放不下就把日志压到 3 行（宁可少几行，也不能出现半行截断）
        for _ in range(4):
            root.update_idletasks()
            if root.winfo_reqheight() <= h_max:
                break
            try:
                cur = int(log_box.cget('height'))
            except Exception:
                break
            if cur <= 3:
                break
            log_box.configure(height=cur - 1)
        # 极矮屏幕：收起顶部副标题与内容区留白，尽量把日志留在可视区
        if root.winfo_reqheight() > h_max:
            try:
                subtitle.pack_forget()
                content.pack_configure(pady=8)
                hrow.pack_configure(pady=(9, 8))
                bar2.pack_configure(pady=8)
                root.update_idletasks()
            except Exception:
                pass
        h = min(max(root.winfo_reqheight() + 10, 640), h_max)
        x = max((sw - w) // 2, 0)
        y = max((shp - h) // 3, 0)
        root.geometry("%dx%d+%d+%d" % (w, h, x, y))
        root.minsize(min(w, root.winfo_reqwidth()), min(h, 600))

    logln("%s v%s 就绪。原始文件全程只读，结果另存为新文件。"
          % (APP_NAME, APP_VER))
    if HAS_DND:
        logln("支持拖拽：把 .xlsx / .xlsm 文件（或整个文件夹）直接拖进窗口即可加入列表。")
    else:
        logln("提示：当前环境缺少拖拽组件（tkdnd），请用「添加文件」按钮选择文件。")
    logln("输出文件夹每次打开都要重新选择；选过的目录会进入历史记录，可从下拉列表直接挑。")

    fit_window()
    pump()
    root.after(1500, lambda: do_check(manual=False))   # 启动静默检查，无新版不打扰

    if ui_check:
        # 自检：把窗口尺寸与底栏控件是否完整可见打成 JSON，供打包/回归验证。
        # XLSX_UI_CHECK_WAIT=6 → 等 6 秒（自动更新检查跑完）再出报告。
        wait_ms = int(float(os.environ.get('XLSX_UI_CHECK_WAIT') or 0) * 1000)

        def ui_report():
            root.update()
            report = {
                'version': APP_VER,
                'window': [root.winfo_width(), root.winfo_height()],
                'requested': [root.winfo_reqwidth(), root.winfo_reqheight()],
                'screen': [root.winfo_screenwidth(), root.winfo_screenheight()],
                'update_button': btn_upd.cget('text'),
                'log_tail': log_box.get('end-6l', 'end').strip().splitlines()[-4:],
                'widgets': {},
            }
            for name, w in (('btn_run', btn_run), ('btn_open', btn_open),
                            ('btn_update', btn_upd), ('btn_exit', btn_exit),
                            ('log_box', log_box), ('listbox', lst)):
                try:
                    report['widgets'][name] = {
                        'y': w.winfo_rooty() - root.winfo_rooty(),
                        'h': w.winfo_height(),
                        'visible': (w.winfo_rooty() - root.winfo_rooty() + w.winfo_height()
                                    <= root.winfo_height()),
                        'text': (w.cget('text') if 'text' in w.keys() else ''),
                    }
                except Exception as exc:                          # noqa: BLE001
                    report['widgets'][name] = {'error': str(exc)}
            print(json.dumps(report, ensure_ascii=False))
            root.destroy()

        if wait_ms > 0:
            root.after(wait_ms, ui_report)
            root.mainloop()
        else:
            ui_report()
        return 0

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
    if args and args[0] == '--check-dnd':
        # 退出码 0 = 拖拽组件可用；3 = 不可用（GUI 仍可运行，但只能用按钮选文件）
        try:
            import tkinterdnd2  # noqa: F401
            print("tkinterdnd2 OK")
            return 0
        except Exception as e:
            sys.stderr.write("tkinterdnd2 不可用：%s\n" % e)
            return 3
    if args and args[0] in ('--check-update', '--check'):
        try:
            info = check_latest()
        except Exception as e:                                  # noqa: BLE001
            print("检查更新失败：%s" % e)
            return 1
        print("当前版本 : v%s" % APP_VER)
        print("最新版本 : v%s%s" % (info['latest'], "" if info['has_update'] else "（已是最新）"))
        if info.get('published'):
            print("发布时间 : %s" % info['published'])
        if info.get('notes'):
            print("更新说明 : %s" % info['notes'].replace("\n", " / "))
        print("清单来源 : %s" % info['source'])
        return 0
    if args and args[0] == '--update':
        if is_frozen():
            print("可执行文件形态请在图形界面里点「检查更新」，命令行不支持自替换。")
            return 1
        try:
            info = check_latest()
            if not info['has_update']:
                print("已是最新版本 v%s。" % APP_VER)
                return 0
            msg, _ = install_update(info, progress=lambda g, t: None)
            print(msg)
            return 0
        except Exception as e:                                  # noqa: BLE001
            print("更新失败：%s" % e)
            return 1
    if args and args[0] == '--version':
        print(APP_VER)
        return 0
    if args and args[0] in ('--help', '-h'):
        print(__doc__)
        return 0
    try:
        return run_gui(ui_check='--ui-check' in args)
    except Exception as e:                      # 没有 tkinter / 无图形环境时退回命令行
        sys.stderr.write("无法启动图形界面：%s\n" % e)
        sys.stderr.write("可用命令行模式：python %s --cli 源文件.xlsx [输出.xlsx]\n"
                         % os.path.basename(__file__))
        return 2


if __name__ == '__main__':
    sys.exit(main())
