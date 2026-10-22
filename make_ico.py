#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从 icon.png 生成 Windows 兼容的 ICO（全部使用 BMP/DIB 条目）。

PIL 的 Image.save(..., format="ICO") 会给每个尺寸写 PNG 压缩条目，
PyInstaller 在 Windows 上嵌入图标时对这类条目容易报错；这里手写标准
BMP(32bpp) 条目，确保 PyInstaller / Windows 资源管理器都能识别。
"""
import io
import os
import struct
import sys

from PIL import Image

SIZES = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128)]


def dib_entry(img):
    """生成 32bpp BMP DIB（含 AND 掩码），返回 (宽, 高, 数据)"""
    w, h = img.size
    px = img.convert("RGBA").load()
    rows = []
    for y in range(h - 1, -1, -1):          # DIB 自下而上
        row = bytearray()
        for x in range(w):
            r, g, b, a = px[x, y]
            row += bytes((b, g, r, a))      # BGRA
        rows.append(bytes(row))
    pixels = b"".join(rows)
    mask_row = ((w + 31) // 32) * 4
    mask = b"\x00" * (mask_row * h)
    header = struct.pack("<IiiHHIIiiII", 40, w, h * 2, 1, 32, 0,
                         len(pixels) + len(mask), 0, 0, 0, 0)
    return w, h, header + pixels + mask


def main(src="icon.png", dst="icon.ico"):
    base = Image.open(src).convert("RGBA")
    entries = []
    for size in SIZES:
        entries.append(dib_entry(base.resize(size, Image.LANCZOS)))
    offset = 6 + 16 * len(entries)
    directory = struct.pack("<HHH", 0, 1, len(entries))
    blobs = b""
    for w, h, data in entries:
        directory += struct.pack("<BBBBHHII", w % 256, h % 256, 0, 0, 1, 32,
                                 len(data), offset + len(blobs))
        blobs += data
    with open(dst, "wb") as f:
        f.write(directory + blobs)

    # 自检
    raw = open(dst, "rb").read()
    res, typ, cnt = struct.unpack("<HHH", raw[:6])
    assert (res, typ) == (0, 1), "ICONDIR 头错误"
    off = 6
    for i in range(cnt):
        w, h, colors, r, planes, bpp, size, ofs = struct.unpack("<BBBBHHII", raw[off:off + 16])
        kind = "PNG" if raw[ofs:ofs + 8] == b"\x89PNG\r\n\x1a\n" else "BMP"
        print("  条目%d: %dx%d %dbpp %dB %s" % (i + 1, w or 256, h or 256, bpp, size, kind))
        off += 16
    print("%s 生成完成，%d 字节（%.1f KB）" % (dst, len(raw), len(raw) / 1024))


if __name__ == "__main__":
    main(*sys.argv[1:])
