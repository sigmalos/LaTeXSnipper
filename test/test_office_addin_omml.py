from __future__ import annotations

from types import SimpleNamespace

import pytest
from lxml import etree

from exporting import formula_converters
from integration.office_addin.omml import MATH, WORD, apply_typography
from integration.office_addin.typography import Typography

OMML = (
    f'<m:oMath xmlns:m="{MATH}" xmlns:w="{WORD}">'
    '<m:r><m:t>12</m:t></m:r>'
    '<m:r><m:t>x</m:t></m:r>'
    '<m:r><m:rPr><m:nor/></m:rPr><m:t>中文</m:t></m:r>'
    '<m:r><w:rPr><w:color w:val="FF0000"/></w:rPr><m:t>y</m:t></m:r>'
    '</m:oMath>'
)


def runs(omml):
    root = etree.fromstring(omml.encode())
    result = []
    for run in root.iter(f"{{{MATH}}}r"):
        properties = run.find(f"{{{WORD}}}rPr")
        fonts = properties.find(f"{{{WORD}}}rFonts")
        result.append({
            "text": "".join(run.itertext()),
            "ascii": fonts.get(f"{{{WORD}}}ascii"),
            "eastAsia": fonts.get(f"{{{WORD}}}eastAsia"),
            "color": properties.find(f"{{{WORD}}}color").get(f"{{{WORD}}}val"),
            "size": properties.find(f"{{{WORD}}}sz").get(f"{{{WORD}}}val"),
        })
    return result


def test_applies_symbol_number_and_cjk_fonts_like_windows():
    typography = Typography(symbol_font_id="mathjax-stix2", number_font_family="Arial",
                            cjk_font_family="Songti SC", font_size_points=10.5, color="#123456")
    result = runs(apply_typography(OMML, typography))
    assert [run["ascii"] for run in result] == ["Arial", "STIX Two Math", "Times New Roman", "STIX Two Math"]
    assert {run["eastAsia"] for run in result} == {"Songti SC"}
    assert {run["size"] for run in result} == {"21"}
    # Source-local color emitted by the converter wins over the default color.
    assert [run["color"] for run in result] == ["123456", "123456", "123456", "FF0000"]


def test_numbers_follow_symbol_font_without_number_family():
    assert runs(apply_typography(OMML, Typography()))[0]["ascii"] == "Cambria Math"


def test_rejects_non_omml_roots():
    with pytest.raises(ValueError):
        apply_typography("<root/>", Typography())


def test_finds_stylesheet_inside_word_for_mac(monkeypatch, tmp_path):
    system, user = tmp_path / "Applications", tmp_path / "home/Applications"
    stylesheet = user / formula_converters._MACOS_WORD_STYLESHEET
    stylesheet.parent.mkdir(parents=True)
    stylesheet.touch()
    monkeypatch.setattr(formula_converters, "sys", SimpleNamespace(platform="darwin"))
    monkeypatch.setattr(formula_converters, "_macos_application_directories", lambda: (system, user))
    assert formula_converters._find_mml2omml_xsl() == stylesheet
    stylesheet.unlink()
    assert formula_converters._find_mml2omml_xsl() is None
