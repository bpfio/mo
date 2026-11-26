#!/usr/bin/env python3
"""docx 兼容性验证器（chengwen skill）

检查四类"本机能过、Word/WPS 拒开"的已知坑（2026-11-26 实测事故沉淀）：
1. zip 完整性 + 全部 XML 部件良构
2. tblW/tcW 宽度必须为整数（五十分之一百分点）——出现 "nn%" 字符串即 FAIL（docx@8.x 指纹）
3. 空 <w:pgNumType/> ——WPS 兼容清单要求剥离（WARN）
4. 页眉页脚 instrText 域白名单（防自定义补丁残留）+ rels 悬挂引用（Word 会弹"修复"）

用法：python3 -I validate_docx.py <file.docx>
退出码：0=通过可投递；1=存在 FAIL
"""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

FIELD_ALLOWLIST = ("PAGE", "PAGEREF", "TOC", "NUMPAGES", "DATE", "SEQ", "STYLEREF")


def main(path):
    fails, warns = [], []
    try:
        z = zipfile.ZipFile(path)
    except Exception as e:
        print(f"FAIL 无法打开 zip: {e}")
        return 1

    corrupt = z.testzip()
    if corrupt:
        fails.append(f"zip 损坏条目: {corrupt}")

    names = set(z.namelist())
    for n in sorted(names):
        if n.endswith((".xml", ".rels")):
            try:
                ET.fromstring(z.read(n))
            except Exception as e:
                fails.append(f"XML 良构 {n}: {e}")

    if "word/document.xml" in names:
        doc = z.read("word/document.xml").decode("utf-8")
        # 宽度整数口径（docx@8.x 输出 "100%" 字符串致 WPS/旧版 Word 拒开）
        for tag in re.findall(r"<w:(?:tbl|tc)W [^>]*/>", doc):
            m = re.search(r'w:w="([^"]*)"', tag)
            if m and not m.group(1).isdigit():
                fails.append(f"宽度非整数口径(docx@8.x指纹): {tag}")
        # 空 pgNumType
        n_empty = doc.count("<w:pgNumType/>")
        if n_empty:
            warns.append(f"空 <w:pgNumType/> ×{n_empty}（WPS 兼容清单要求剥离）")

    # 页眉页脚域白名单
    for n in sorted(names):
        if re.match(r"word/(footer|header)\d*\.xml$", n):
            xml = z.read(n).decode("utf-8")
            for instr in re.findall(r"<w:instrText[^>]*>([^<]*)</w:instrText>", xml):
                s = instr.strip()
                if s and not s.startswith(FIELD_ALLOWLIST):
                    fails.append(f"{n} 含白名单外域: {s!r}")

    # document.xml.rels 悬挂引用（Word 弹"修复"对话框的常见诱因）
    rels_name = "word/_rels/document.xml.rels"
    if rels_name in names:
        rels = z.read(rels_name).decode("utf-8")
        for target in re.findall(r'Target="([^"]+)"[^>]*?(?:TargetMode="External")?', rels):
            if target.startswith(("http", "mailto", "file:")):
                continue
            full = "word/" + target if not target.startswith("/") else target[1:]
            if full not in names:
                fails.append(f"rels 悬挂引用: {target} -> {full} 不存在")

    for f in fails:
        print(f"FAIL {f}")
    for w in warns:
        print(f"WARN {w}")
    if fails:
        print(f"\n结论：不通过（{len(fails)} 项 FAIL），禁止投递")
        return 1
    print(f"\n结论：通过（0 FAIL，{len(warns)} WARN）——可投递")
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(sys.argv[1]))
