#!/usr/bin/env python3
"""
doc2md.py - 万物转 Markdown，支持 docx（含图片）、PDF、PPTX、XLSX、图片、HTML、ZIP
用法：
    python3 doc2md.py 文件.docx
    python3 doc2md.py 文件.pdf
    python3 doc2md.py 文件夹/          # 批量处理文件夹内所有支持的文件
"""

import sys
import base64
import argparse
from pathlib import Path


def convert_docx(src: Path) -> str:
    """docx 转 Markdown，图片嵌入为 base64"""
    from docx import Document
    from docx.oxml.ns import qn

    doc = Document(str(src))

    # 建立 rId -> base64 data URI 的映射
    image_map = {}
    for rId, rel in doc.part.rels.items():
        if "image" in rel.reltype:
            blob = rel.target_part.blob
            ctype = rel.target_part.content_type
            b64 = base64.b64encode(blob).decode()
            image_map[rId] = f"data:{ctype};base64,{b64}"

    img_counter = [0]

    def para_to_md(para) -> str:
        result = []
        for child in para._p:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "r":
                for blip in child.findall(".//" + qn("a:blip")):
                    rId = blip.get(
                        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
                    )
                    if rId and rId in image_map:
                        img_counter[0] += 1
                        result.append(f"\n![图{img_counter[0]}]({image_map[rId]})\n")
                for t in child.findall(qn("w:t")):
                    if t.text:
                        result.append(t.text)
            elif tag == "hyperlink":
                for t in child.findall(".//" + qn("w:t")):
                    if t.text:
                        result.append(t.text)
        return "".join(result)

    lines = []
    for para in doc.paragraphs:
        style = para.style.name
        text = para_to_md(para)
        if not text.strip():
            continue
        if style.startswith("Heading 1"):
            lines.append(f"# {text}")
        elif style.startswith("Heading 2"):
            lines.append(f"## {text}")
        elif style.startswith("Heading 3"):
            lines.append(f"### {text}")
        else:
            lines.append(text)

    print(f"  [docx] {img_counter[0]} 张图片，{len(lines)} 个段落", file=sys.stderr)
    return "\n\n".join(lines)


def convert_other(src: Path) -> str:
    """其他格式交给 markitdown 处理"""
    from markitdown import MarkItDown
    md = MarkItDown()
    result = md.convert(str(src))
    return result.text_content


SUPPORTED = {
    ".docx": convert_docx,
    ".pdf": convert_other,
    ".pptx": convert_other,
    ".xlsx": convert_other,
    ".xls": convert_other,
    ".jpg": convert_other,
    ".jpeg": convert_other,
    ".png": convert_other,
    ".gif": convert_other,
    ".webp": convert_other,
    ".html": convert_other,
    ".htm": convert_other,
    ".zip": convert_other,
    ".mp3": convert_other,
    ".wav": convert_other,
    ".m4a": convert_other,
}


def process_file(src: Path, output_dir: Path = None) -> Path:
    ext = src.suffix.lower()
    if ext not in SUPPORTED:
        print(f"跳过不支持的格式：{src.name}", file=sys.stderr)
        return None

    print(f"转换：{src.name}", file=sys.stderr)
    content = SUPPORTED[ext](src)

    out_dir = output_dir or src.parent
    out_path = out_dir / (src.stem + ".md")
    out_path.write_text(content, encoding="utf-8")
    print(f"  → {out_path}", file=sys.stderr)
    return out_path


def main():
    parser = argparse.ArgumentParser(description="万物转 Markdown")
    parser.add_argument("inputs", nargs="+", help="文件或文件夹路径")
    parser.add_argument("-o", "--output-dir", help="输出目录（默认与源文件同目录）")
    args = parser.parse_args()

    output_dir = Path(args.output_dir) if args.output_dir else None
    if output_dir:
        output_dir.mkdir(parents=True, exist_ok=True)

    for inp in args.inputs:
        p = Path(inp)
        if p.is_dir():
            files = [f for f in p.rglob("*") if f.suffix.lower() in SUPPORTED]
            print(f"文件夹 {p}：找到 {len(files)} 个文件", file=sys.stderr)
            for f in files:
                process_file(f, output_dir)
        elif p.is_file():
            process_file(p, output_dir)
        else:
            print(f"路径不存在：{inp}", file=sys.stderr)


if __name__ == "__main__":
    main()


# # 单文件
# python3 doc2md.py 文件.docx
# python3 doc2md.py 文件.pdf

# # 指定输出目录
# python3 doc2md.py 文件.pptx -o ./output/

# # 批量处理整个文件夹
# python3 doc2md.py 资料文件夹/

# python3 doc2md.py /Users/chuangzhidian/Documents/我用Hermes搞定了微信文章稳定提取方案.docx -o ./output/微信文章稳定提取方案.md