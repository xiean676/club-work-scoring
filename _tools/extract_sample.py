#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""锚点样本提取器：把 docx / pptx / pdf 的文本层抽出来，供人工研读。
输出到 _work/samples/<name>.txt，纯标准库实现。
"""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "_work" / "samples"
OUT.mkdir(parents=True, exist_ok=True)


def _paras(xml_bytes):
    root = ET.fromstring(xml_bytes)
    buf = []
    for p in root.iter():
        if p.tag.split("}")[-1] != "p":
            continue
        line = "".join(t.text or "" for t in p.iter() if t.tag.split("}")[-1] == "t")
        if line.strip():
            buf.append(line.strip())
    return buf


def docx_text(path):
    parts = []
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        for name in ["word/document.xml"] + sorted(n for n in names if re.search(r"word/(header|footer)\d*\.xml", n)):
            if name not in names:
                continue
            try:
                txt = "\n".join(_paras(zf.read(name)))
            except ET.ParseError:
                continue
            if txt.strip():
                parts.append("[%s]\n%s" % (name, txt))
    return "\n\n".join(parts)


def pptx_text(path):
    with zipfile.ZipFile(path) as zf:
        slides = sorted(
            (n for n in zf.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
            key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)),
        )
        out = ["[共 %d 页]" % len(slides)]
        for name in slides:
            idx = int(re.search(r"(\d+)", name.split("/")[-1]).group(1))
            try:
                lines = _paras(zf.read(name))
            except ET.ParseError:
                lines = []
            notes = ""
            nname = "ppt/notesSlides/notesSlide%d.xml" % idx
            if nname in zf.namelist():
                try:
                    nlines = [x for x in _paras(zf.read(nname)) if not re.fullmatch(r"\d+", x)]
                    notes = "\n".join(nlines)
                except ET.ParseError:
                    notes = ""
            out.append("\n===== 第 %d 页 =====\n%s" % (idx, "\n".join(lines)))
            if notes.strip():
                out.append("[备注]\n%s" % notes)
        media = [n for n in zf.namelist() if n.startswith("ppt/media/")]
        out.append("[媒体文件 %d 个] %s" % (len(media), ", ".join(Path(m).name for m in media[:40])))
        return "\n\n".join(out)


def docx_stats(path):
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        media = [n for n in names if n.startswith("word/media/")]
        core = zf.read("docProps/core.xml").decode("utf-8", "ignore") if "docProps/core.xml" in names else ""
        app = zf.read("docProps/app.xml").decode("utf-8", "ignore") if "docProps/app.xml" in names else ""
    meta = {}
    for blob in (core, app):
        for m in re.finditer(r"<(\w+)[^>]*>([^<]{1,80})</\1>", blob):
            meta.setdefault(m.group(1), m.group(2))
    return media, meta


def main():
    for raw in sys.argv[1:]:
        p = Path(raw)
        if not p.exists():
            print("缺失：%s" % p)
            continue
        name = p.stem
        if p.suffix.lower() == ".docx":
            text = docx_text(p)
            media, meta = docx_stats(p)
            text += "\n\n[内嵌媒体 %d 个：%s]\n[文档属性] %s" % (
                len(media), ", ".join(Path(m).name for m in media[:30]), meta)
        elif p.suffix.lower() == ".pptx":
            text = pptx_text(p)
        else:
            print("跳过 %s" % p)
            continue
        dest = OUT / (re.sub(r'[\\/:*?"<>|]', "_", name) + ".txt")
        dest.write_text(text, encoding="utf-8")
        print("%-28s -> %s  (%d 字)" % (name, dest.name, len(text)))


if __name__ == "__main__":
    main()
