"""嵌入 xlsx 编辑测试。"""
import io
import zipfile

import pytest
from openpyxl import load_workbook

from docxsurgeon import set_table_range

EMBED = "word/embeddings/data.xlsx"


def _read_table_ref(pkg):
    """从包内的嵌入 xlsx 字节里直接读 xl/tables/table1.xml 的 ref 属性。"""
    import re
    with zipfile.ZipFile(io.BytesIO(pkg.read_part(EMBED))) as zf:
        xml = zf.read("xl/tables/table1.xml").decode("utf-8")
    m = re.search(r'ref="([^"]+)"', xml)
    return m.group(1)


def test_open_embedded_xlsx_write(pkg):
    with pkg.open_embedded_xlsx(EMBED) as wb:
        ws = wb.active
        ws["A2"] = "9月"
        ws["B2"] = 42
        set_table_range(ws, "A1:B8")

    # 单元格修改已写回
    with pkg.open_embedded_xlsx(EMBED) as wb:
        ws = wb.active
        assert ws["A2"].value == "9月"
        assert ws["B2"].value == 42
    # 表格范围已写回（直接检查嵌套 zip 的 XML）
    assert _read_table_ref(pkg) == "A1:B8"


def test_open_embedded_xlsx_exception_no_write(pkg):
    original = pkg.read_part(EMBED)
    with pytest.raises(RuntimeError):
        with pkg.open_embedded_xlsx(EMBED) as wb:
            wb.active["A2"] = "不该写进去"
            raise RuntimeError("boom")
    assert pkg.read_part(EMBED) == original, "异常退出时不应写回修改"


def test_set_table_range_multiple_tables(pkg):
    with pkg.open_embedded_xlsx(EMBED) as wb:
        ws = wb.active
        n = set_table_range(ws, "A1:B10")
        assert n == 1


def test_set_table_range_no_table_raises():
    from openpyxl import Workbook
    ws = Workbook().active
    with pytest.raises(ValueError):
        set_table_range(ws, "A1:B5")


def test_open_embedded_xlsx_unknown_raises(pkg):
    with pytest.raises(KeyError):
        with pkg.open_embedded_xlsx("word/embeddings/nothere.xlsx"):
            pass


def test_roundtrip_through_openpyxl_keeps_table(pkg, tmp_path):
    """openpyxl 往返后嵌入文件仍是合法 xlsx 且表格仍在。"""
    with pkg.open_embedded_xlsx(EMBED) as wb:
        wb.active["C1"] = "新增列"
    out = tmp_path / "roundtrip.xlsx"
    out.write_bytes(pkg.read_part(EMBED))
    wb2 = load_workbook(out)
    assert wb2.active["C1"].value == "新增列"
    assert len(wb2.active.tables) == 1, "表格在 openpyxl 往返后丢失"
