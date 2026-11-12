"""MathML to Word OMML with LaTeXSnipper typography.

Ports ``OmmlTypographyMapper`` from the Windows plugin: Word owns equation
layout, while the formula's symbol, number and CJK fonts, size and color are
applied as run properties that local source styles still override.
"""

from __future__ import annotations

import re

from integration.office_addin.typography import Typography

MATH = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WORD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
SYMBOL_FAMILIES = {"mathjax-tex": "Cambria Math", "mathjax-stix2": "STIX Two Math"}
_NUMBER = re.compile(r"\A\d+(?:[.,]\d+)*\Z")


def convert(mathml: str, typography: Typography) -> str:
    from exporting.formula_converters import mathml_to_omml

    return apply_typography(mathml_to_omml(mathml), typography)


def apply_typography(omml: str, typography: Typography) -> str:
    from lxml import etree

    parser = etree.XMLParser(resolve_entities=False, no_network=True, load_dtd=False)
    root = etree.fromstring(omml.encode("utf-8"), parser)
    if root.tag not in {f"{{{MATH}}}oMath", f"{{{MATH}}}oMathPara"}:
        raise ValueError("需要 OMML 公式根节点。")
    symbol = SYMBOL_FAMILIES.get(typography.symbol_font_id)
    if symbol is None:
        raise ValueError("不支持的数学字体方案。")
    size = str(int(typography.font_size_points * 2 + 0.5))
    color = typography.color[1:]
    for run in root.iter(f"{{{MATH}}}r"):
        text = "".join(node.text or "" for node in run.findall(f"{{{MATH}}}t"))
        normal = run.find(f"{{{MATH}}}rPr/{{{MATH}}}nor")
        value = None if normal is None else normal.get(f"{{{MATH}}}val")
        is_text = normal is not None and value not in {"0", "false"}
        is_number = not is_text and bool(_NUMBER.match(text))
        family = (
            "Times New Roman"
            if is_text
            else typography.number_font_family
            if is_number and typography.number_font_family
            else symbol
        )
        properties = run.find(f"{{{WORD}}}rPr")
        if properties is None:
            properties = etree.Element(f"{{{WORD}}}rPr", nsmap={"w": WORD})
            math_properties = run.find(f"{{{MATH}}}rPr")
            if math_properties is None:
                run.insert(0, properties)
            else:
                math_properties.addnext(properties)
        fonts = properties.find(f"{{{WORD}}}rFonts")
        if fonts is None:
            fonts = etree.Element(f"{{{WORD}}}rFonts", nsmap={"w": WORD})
            properties.insert(0, fonts)
        fonts.set(f"{{{WORD}}}ascii", family)
        fonts.set(f"{{{WORD}}}hAnsi", family)
        fonts.set(f"{{{WORD}}}eastAsia", typography.cjk_font_family)
        for theme in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme"):
            fonts.attrib.pop(f"{{{WORD}}}{theme}", None)
        for name, setting in (("color", color), ("sz", size), ("szCs", size)):
            if properties.find(f"{{{WORD}}}{name}") is None:
                element = etree.SubElement(properties, f"{{{WORD}}}{name}")
                element.set(f"{{{WORD}}}val", setting)
    return etree.tostring(root, encoding="unicode")
