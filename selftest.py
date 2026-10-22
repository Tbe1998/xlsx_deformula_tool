#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""自检脚本（不依赖第三方库，CI 与本地都能跑）。

自己造一个最小 xlsx（带公式 + 缓存结果，模拟 Excel/WPS 保存过的文件），
交给 xlsx_deformula.convert() 处理，然后断言：
  1. 输出文件里没有残留的普通公式；
  2. 每个单元格的值与原缓存结果一致；
  3. 源文件字节级未被改动；
  4. 内嵌图片单元格 (=DISPIMG) 的公式被原样保留；
  5. 拖拽路径解析（多文件、带空格、整个文件夹）与输出目录历史记忆正常。

用法：python selftest.py
"""
import hashlib
import os
import re
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import xlsx_deformula as tool  # noqa: E402

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/cellimages.xml" ContentType="application/vnd.wps-officedocument.cellimage+xml"/>
</Types>'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

WB = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="数据" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/></workbook>'''

WB_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>'''

SHEET = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>3</v></c>
<c r="C1"><f>A1+B1</f><v>5</v></c>
<c r="D1" t="str"><f>"编号-"&amp;A1</f><v>编号-2</v></c></row>
<row r="2"><c r="A2" t="str"><f>_xlfn.DISPIMG("ID_TESTIMAGE",1)</f><v>=DISPIMG("ID_TESTIMAGE",1)</v></c></row>
</sheetData></worksheet>'''

CELLIMG = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<etc:cellImages xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData">
<etc:cellImage><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="4" name="ID_TESTIMAGE" descr="测试图"/></xdr:nvPicPr>
<xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic></etc:cellImage></etc:cellImages>'''

CELLIMG_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>'''

CALCCHAIN = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="C1" i="1"/></calcChain>'''

PNG_1PX = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4"
    "890000000a49444154789c6360000002000100ffff03000006000557bfabd400"
    "00000049454e44ae426082")


def build_sample(path):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CT)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('xl/workbook.xml', WB)
        z.writestr('xl/_rels/workbook.xml.rels', WB_RELS)
        z.writestr('xl/worksheets/sheet1.xml', SHEET)
        z.writestr('xl/cellimages.xml', CELLIMG)
        z.writestr('xl/_rels/cellimages.xml.rels', CELLIMG_RELS)
        z.writestr('xl/media/image1.png', PNG_1PX)
        z.writestr('xl/calcChain.xml', CALCCHAIN)


def md5(path):
    with open(path, 'rb') as f:
        return hashlib.md5(f.read()).hexdigest()


def main():
    tmp = tempfile.mkdtemp(prefix='xlsx_deformula_test_')
    src = os.path.join(tmp, '示例.xlsx')
    dst = os.path.join(tmp, '示例_无公式.xlsx')
    build_sample(src)
    before = md5(src)

    pre = tool.scan_workbook(src)
    stats = tool.convert(src, dst)
    fails = []

    # 1) 源文件未被改动
    if md5(src) != before:
        fails.append("源文件被修改了")

    # 2) 公式残留（DISPIMG 除外）
    with zipfile.ZipFile(dst) as z:
        xml = z.read('xl/worksheets/sheet1.xml').decode()
        blocks = re.findall(r'<c\b[^>]*?(?:/>|>.*?</c>)', xml, re.S)
        left = [b for b in blocks if '<f' in b and 'DISPIMG' not in b]
        if left:
            fails.append("仍有普通公式残留: %s" % left)
        # 3) 值一致
        def cell(coord):
            for b in blocks:
                if re.match(r'<c r="%s"' % coord, b):
                    return b
            return ''
        if '<v>5</v>' not in cell('C1'):
            fails.append("C1 计算值丢失: %s" % cell('C1'))
        if '编号-2' not in cell('D1'):
            fails.append("D1 文本结果丢失: %s" % cell('D1'))
        # 4) DISPIMG 单元格保留
        if 'DISPIMG' not in cell('A2'):
            fails.append("内嵌图片单元格未保留: %s" % cell('A2'))
        # 5) calcChain 移除
        if 'xl/calcChain.xml' in z.namelist():
            fails.append("calcChain.xml 未移除")

    # 拖拽解析 / 历史记忆
    fails.extend(test_dropped_paths_and_history())

    # 输出路径保护：不允许覆盖源文件
    try:
        tool.convert(src, src)
        fails.append("未阻止把结果写回源文件")
    except ValueError:
        pass

    print("预检: %s" % pre)
    print("统计: %s" % stats)
    print("临时目录: %s" % tmp)
    if fails:
        print("\n自检失败：")
        for f in fails:
            print("  ✘ " + f)
        return 1
    print("\n自检通过：公式已转为数值、值一致、图片单元格保留、源文件未改动。")
    return 0


def test_dropped_paths_and_history():
    """拖拽路径解析 + 输出目录历史记忆 的单元测试。"""
    fails = []
    d = tempfile.mkdtemp(prefix='xlsx_deformula_drop_')
    for name in ('a.xlsx', 'b.XLSM', 'c.txt', '~$lock.xlsx'):
        open(os.path.join(d, name), 'w').write('x')
    spaced = os.path.join(tempfile.mkdtemp(prefix='xlsx_deformula_sp_'), '空 格 文件.xlsx')
    open(spaced, 'w').write('x')

    # 多文件 + 带空格（Tk 用花括号包住）+ 非表格条目
    found, ignored = tool.parse_dropped_paths('{%s} %s %s'
                                              % (spaced, os.path.join(d, 'a.xlsx'),
                                                 os.path.join(d, 'c.txt')))
    names = [os.path.basename(x) for x in found]
    if '空 格 文件.xlsx' not in names:
        fails.append("带空格路径解析失败: %s" % names)
    if 'a.xlsx' not in names:
        fails.append("普通路径解析失败: %s" % names)
    if not any('c.txt' == os.path.basename(x) for x in ignored):
        fails.append("非 Excel 条目未被忽略")

    # 拖入整个文件夹：只收表格，跳过 txt 与 ~$ 临时文件
    found2, _ = tool.parse_dropped_paths(d)
    names2 = sorted(os.path.basename(x) for x in found2)
    if names2 != ['a.xlsx', 'b.XLSM']:
        fails.append("文件夹解析结果不符: %s" % names2)

    # 历史记忆：最近优先、去重、最多 10 条
    cfg_home = tempfile.mkdtemp(prefix='xlsx_deformula_cfg_')
    old_env = os.environ.get('XDG_CONFIG_HOME'), os.environ.get('APPDATA')
    os.environ['XDG_CONFIG_HOME'] = cfg_home
    os.environ['APPDATA'] = cfg_home
    try:
        for i in range(1, 13):
            tool.remember_output_dir(os.path.join(cfg_home, 'out%d' % i))
        tool.remember_output_dir(os.path.join(cfg_home, 'out5'))      # 重复项应提前
        hist = tool.load_config()['output_history']
        if len(hist) != 10:
            fails.append("历史条数应为 10，实际 %d" % len(hist))
        if not hist[0].endswith('out5'):
            fails.append("最近使用的目录没有排在最前: %s" % hist[0])
        if sum(1 for h in hist if h.endswith('out5')) != 1:
            fails.append("历史没有去重")
        if not os.path.exists(tool.config_file()):
            fails.append("配置文件未落盘: %s" % tool.config_file())
    finally:
        for k, v in zip(('XDG_CONFIG_HOME', 'APPDATA'), old_env):
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v

    for f in fails:
        print("  ✘ " + f)
    print("拖拽解析 / 历史记忆 测试：%s" % ("失败" if fails else "通过"))
    return fails


def verify_output(path):
    """校验一个处理结果文件（供 CI 校验打包后 exe 的产物）。"""
    fails = []
    with zipfile.ZipFile(path) as z:
        xml = z.read('xl/worksheets/sheet1.xml').decode()
        blocks = re.findall(r'<c\b[^>]*?(?:/>|>.*?</c>)', xml, re.S)
        left = [b for b in blocks if '<f' in b and 'DISPIMG' not in b]
        if left:
            fails.append("仍有普通公式残留: %s" % left)
        if not any('<v>5</v>' in b for b in blocks):
            fails.append("C1 计算值 5 丢失")
        if not any('编号-2' in b for b in blocks):
            fails.append("D1 文本结果丢失")
        if not any('DISPIMG' in b for b in blocks):
            fails.append("内嵌图片单元格未保留")
        if 'xl/calcChain.xml' in z.namelist():
            fails.append("calcChain.xml 未移除")
    for f in fails:
        print("  ✘ " + f)
    print("产物校验：%s" % ("失败" if fails else "通过"))
    return 1 if fails else 0


def main_argv(argv):
    if argv[:1] == ['--make-sample']:
        build_sample(argv[1])
        print("已生成样例: %s" % argv[1])
        return 0
    if argv[:1] == ['--verify']:
        return verify_output(argv[1])
    return main()


if __name__ == '__main__':
    sys.exit(main_argv(sys.argv[1:]))
