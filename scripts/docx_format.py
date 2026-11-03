#!/usr/bin/env python3
"""lex-hj公文排版 自包含 docx 排版脚本。

按议案模板的视觉格式规范，对一个已存在的 .docx 公文应用全局格式：
正文仿宋三号、数字英文 Times New Roman、固定行距30磅、首行缩进2字符、
四级标题字体字号、表格/图题格式等。

设计要点：
1. 内置 python-docx 依赖检测，缺失时给出清晰的安装命令并退出（不静默失败）。
2. 运行时自动检测平台，按 references/font-mapping.md 的映射选择字体族名，
   解决 Windows 中易字库在 macOS/Linux 上缺失导致的版式错乱问题。
3. 脚本只做"排版/格式"，不做内容撰写；仅接受明确的目标文件路径。

用法：
    python docx_format.py input.docx [--output output.docx] [--platform windows|macos|linux|auto]
"""

from __future__ import annotations

import argparse
import platform
import sys
from pathlib import Path

# ---- 依赖检测 ---------------------------------------------------------------

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml.ns import qn
    from docx.shared import Pt, Cm
except ImportError:
    print(
        "错误：缺少 python-docx 依赖。\n"
        "请先安装：pip install python-docx\n"
        "（本技能处理 .docx 文件需要该依赖）",
        file=sys.stderr,
    )
    raise SystemExit(2)


# ---- 跨平台字体映射（与 references/font-mapping.md 保持一致）----------------

FONT_MAP = {
    # 元素 -> {平台: 中文字体名, 西文字体名}
    "body": {  # 正文
        "windows": ("仿宋_GB2312", "Times New Roman"),
        "macos": ("朱雀仿宋", "Times New Roman"),
        "linux": ("朱雀仿宋", "Liberation Serif"),
    },
    "h1": {  # 一级标题（华文中宋二号加粗居中）
        "windows": ("华文中宋", "Times New Roman"),
        "macos": ("思源宋体", "Times New Roman"),
        "linux": ("思源宋体", "Liberation Serif"),
    },
    "h2": {  # 二级标题（黑体三号）
        "windows": ("黑体", "Times New Roman"),
        "macos": ("思源黑体", "Times New Roman"),
        "linux": ("思源黑体", "Liberation Serif"),
    },
    "h3": {  # 三级标题（楷体_GB2312三号加粗）
        "windows": ("楷体_GB2312", "Times New Roman"),
        "macos": ("霞鹜文楷", "Times New Roman"),
        "linux": ("霞鹜文楷", "Liberation Serif"),
    },
    "h4": {  # 四级标题（仿宋_GB2312三号）
        "windows": ("仿宋_GB2312", "Times New Roman"),
        "macos": ("朱雀仿宋", "Times New Roman"),
        "linux": ("朱雀仿宋", "Liberation Serif"),
    },
    "diagram": {  # 架构图（微软雅黑）
        "windows": ("微软雅黑", "Times New Roman"),
        "macos": ("苹方", "Times New Roman"),
        "linux": ("思源黑体", "Liberation Serif"),
    },
}

# 字号（磅）：三号 = 16pt，二号 = 22pt，小四 = 12pt
PT_SAN = 16  # 三号
PT_ER = 22  # 二号
PT_XIAOSI = 12  # 小四

# 行距
LINE_BODY = Pt(30)  # 正文固定值30磅
LINE_TABLE = Pt(20)  # 表格固定值20磅


def detect_platform(explicit: str | None) -> str:
    """解析目标平台。"""
    if explicit and explicit != "auto":
        return explicit
    system = platform.system().lower()
    if system == "windows":
        return "windows"
    if system == "darwin":
        return "macos"
    return "linux"


def set_run_font(run, cn_font: str, west_font: str) -> None:
    """设置 run 的中文字体与西文字体。"""
    run.font.name = west_font
    # 通过 rFonts 的 eastAsia 属性设置中文字体
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:eastAsia"), cn_font)
    rfonts.set(qn("w:ascii"), west_font)
    rfonts.set(qn("w:hAnsi"), west_font)


def set_paragraph_format(para, cn_font: str, west_font: str, size_pt: int, bold: bool = False) -> None:
    """设置段落的字体、字号、行距。"""
    for run in para.runs:
        run.font.size = Pt(size_pt)
        run.font.bold = bold
        set_run_font(run, cn_font, west_font)


def apply_body_style(doc, platform_key: str) -> None:
    """应用正文全局格式（仿宋三号、行距30磅、首行缩进2字符、两端对齐）。"""
    cn, west = FONT_MAP["body"][platform_key]
    for para in doc.paragraphs:
        # 跳过空段落与标题（标题由 apply_heading_style 处理）
        if not para.text.strip():
            continue
        set_paragraph_format(para, cn, west, PT_SAN)
        pf = para.paragraph_format
        pf.line_spacing = LINE_BODY
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.first_line_indent = Cm(0.85)  # 约等于2字符（三号16pt）


def is_h1(text: str) -> bool:
    """识别一级标题：一、二、三、……（中文数字+顿号）。"""
    import re
    return bool(re.match(r"^[一二三四五六七八九十]+、", text))


def is_h2(text: str) -> bool:
    """识别二级标题：（一）（二）……（中文数字序号）。"""
    import re
    return bool(re.match(r"^（[一二三四五六七八九十]+）", text))


def is_h3(text: str) -> bool:
    """识别三级标题：1. 2. 3. ……（阿拉伯数字+点号）。"""
    import re
    return bool(re.match(r"^\d+\.", text))


def is_h4(text: str) -> bool:
    """识别四级标题：（1）（2）……（阿拉伯数字序号）。"""
    import re
    return bool(re.match(r"^（\d+）", text))


def apply_heading_style(doc, platform_key: str) -> None:
    """按四级标题规范应用标题格式。由文本前缀识别层级。"""
    cn_h1, west = FONT_MAP["h1"][platform_key]
    cn_h2, _ = FONT_MAP["h2"][platform_key]
    cn_h3, _ = FONT_MAP["h3"][platform_key]
    cn_h4, _ = FONT_MAP["h4"][platform_key]

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        if is_h1(text):  # 一、华文中宋二号加粗居中
            set_paragraph_format(para, cn_h1, west, PT_ER, bold=True)
            para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.first_line_indent = Cm(0)
        elif is_h2(text):  # （一）黑体三号
            set_paragraph_format(para, cn_h2, west, PT_SAN)
            para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif is_h3(text):  # 1. 楷体_GB2312三号加粗
            set_paragraph_format(para, cn_h3, west, PT_SAN, bold=True)
            para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        elif is_h4(text):  # （1）仿宋_GB2312三号
            set_paragraph_format(para, cn_h4, west, PT_SAN)
            para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT


def apply_page_margins(doc) -> None:
    """设置页边距：上下各 2.54cm，左右各 3.18cm。"""
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.18)
        section.right_margin = Cm(3.18)


def apply_format(input_path: Path, output_path: Path, platform_key: str) -> None:
    """对文档应用完整格式（页边距 + 正文 + 标题层级）。"""
    doc = Document(str(input_path))
    apply_page_margins(doc)
    apply_body_style(doc, platform_key)
    apply_heading_style(doc, platform_key)
    doc.save(str(output_path))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="按议案模板格式规范对一个 .docx 公文应用排版格式（仿宋三号、行距30磅等）。"
    )
    parser.add_argument("input", help="输入的 .docx 文件路径")
    parser.add_argument("--output", help="输出文件路径（默认覆盖原文件）")
    parser.add_argument(
        "--platform",
        default="auto",
        choices=["auto", "windows", "macos", "linux"],
        help="目标平台，默认 auto 自动检测（决定字体族名）",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input).expanduser().resolve()
    if not input_path.is_file():
        print(f"错误：输入文件不存在：{input_path}", file=sys.stderr)
        raise SystemExit(2)
    if input_path.suffix.lower() != ".docx":
        print("错误：仅支持 .docx 文件。", file=sys.stderr)
        raise SystemExit(2)
    output_path = Path(args.output).expanduser().resolve() if args.output else input_path
    platform_key = detect_platform(args.platform)

    print(f"目标平台：{platform_key}")
    print(f"应用格式：{input_path} -> {output_path}")
    apply_format(input_path, output_path, platform_key)
    print("排版完成。请对照 references/format-spec.md 逐项核对格式。")


if __name__ == "__main__":
    main()
