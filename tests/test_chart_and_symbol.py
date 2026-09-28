"""图表引用范围与复选框符号测试。"""
import zipfile

import pytest

from docxsurgeon import DocxPackage

CHART = "word/charts/chart1.xml"

CHART_XML_MULTI_COLUMN = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">'
    '<c:chart><c:plotArea><c:lineChart><c:ser>'
    '<c:val><c:numRef><c:f>Sheet1!$AB$2:$AB$20</c:f></c:numRef></c:val>'
    '<c:val><c:numRef><c:f>Sheet1!$AC$2:$AC$20</c:f></c:numRef></c:val>'
    '</c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>'
)


def test_set_chart_data_rows_basic(pkg):
    n = pkg.set_chart_data_rows(CHART, last_row=8)
    assert n == 2  # cat + val 两处引用
    text = pkg.read_part(CHART).decode("utf-8")
    assert "Sheet1!$A$2:$A$8" in text
    assert "Sheet1!$B$2:$B$8" in text
    assert "Sheet1!$A$2:$A$20" not in text
    assert "Sheet1!$B$2:$B$20" not in text
    # 系列名引用（$B$1，单行）不受影响
    assert "Sheet1!$B$1" in text


def test_set_chart_data_rows_shrink_to_single(pkg):
    n = pkg.set_chart_data_rows(CHART, last_row=2)
    assert n == 2
    text = pkg.read_part(CHART).decode("utf-8")
    assert "Sheet1!$A$2" in text
    assert "$A$2:$A$" not in text, "last_row<=start_row 时应收缩为单行引用"


def test_set_chart_data_rows_multi_letter_columns(tmp_path):
    path = tmp_path / "multi.docx"
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("word/charts/chart1.xml", CHART_XML_MULTI_COLUMN)
    pkg = DocxPackage(path)
    n = pkg.set_chart_data_rows("word/charts/chart1.xml", last_row=12)
    assert n == 2
    text = pkg.read_part("word/charts/chart1.xml").decode("utf-8")
    assert "Sheet1!$AB$2:$AB$12" in text
    assert "Sheet1!$AC$2:$AC$12" in text


def test_set_chart_data_rows_other_sheet_untouched(tmp_path):
    path = tmp_path / "sheet.docx"
    xml = ('<c:chartSpace xmlns:c="http://x">'
           '<c:f>Sheet1!$B$2:$B$20</c:f>'
           '<c:f>数据!$B$2:$B$20</c:f>'
           '</c:chartSpace>')
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("word/charts/chart1.xml", xml)
    pkg = DocxPackage(path)
    n = pkg.set_chart_data_rows("word/charts/chart1.xml", last_row=6)
    assert n == 1, "只应改写指定 sheet 的引用"
    text = pkg.read_part("word/charts/chart1.xml").decode("utf-8")
    assert "Sheet1!$B$2:$B$6" in text
    assert "数据!$B$2:$B$20" in text


def test_set_chart_data_rows_no_match(pkg):
    n = pkg.set_chart_data_rows("word/charts/chart1.xml", last_row=8, sheet="不存在的表")
    assert n == 0


def test_fix_checkbox_glyph(pkg):
    n = pkg.fix_checkbox_glyph()
    assert n == 2, "文档中有两个 ☑"
    text = pkg.read_part("word/document.xml").decode("utf-8")
    assert "☑" not in text
    assert '<w:sym w:font="Wingdings 2" w:char="0052"/>' in text
    # 符号 run 前后正确闭合了文本 run
    assert text.count("</w:t></w:r><w:r><w:rPr>") == 2
    assert text.count('<w:t xml:space="preserve">') == 2
    # 没有动其他部件
    assert pkg.has_part("word/embeddings/data.xlsx")


def test_fix_checkbox_glyph_custom_symbol(pkg):
    pkg.fix_checkbox_glyph(symbol_font="Wingdings", symbol_char="00FC")
    text = pkg.read_part("word/document.xml").decode("utf-8")
    assert 'w:font="Wingdings" w:char="00FC"' in text
