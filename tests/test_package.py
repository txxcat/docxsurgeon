"""部件级操作与保存保真测试。"""
import zipfile

from docxsurgeon import DocxPackage


def test_unmodified_parts_byte_identical(docx_path, tmp_path):
    """保存后，未修改的部件必须字节级原样保留。"""
    pkg = DocxPackage(docx_path)
    out = tmp_path / "out.docx"
    pkg.save(out)

    with zipfile.ZipFile(docx_path) as a, zipfile.ZipFile(out) as b:
        names_a = a.namelist()
        assert b.namelist() == names_a, "部件顺序或集合发生变化"
        for name in names_a:
            assert a.read(name) == b.read(name), "部件 %r 内容被意外改动" % name


def test_list_and_read_parts(pkg):
    parts = pkg.list_parts()
    assert "word/document.xml" in parts
    assert "word/embeddings/data.xlsx" in parts
    assert pkg.has_part("word/charts/chart1.xml")
    assert not pkg.has_part("word/nothing.xml")
    assert pkg.read_part("[Content_Types].xml") == b"<Types/>"


def test_replace_part(pkg):
    pkg.replace_part("word/document.xml", "<new/>")
    assert pkg.read_part("word/document.xml") == b"<new/>"
    # 其他部件不受影响
    assert pkg.has_part("word/charts/chart1.xml")


def test_replace_part_unknown_raises(pkg):
    import pytest
    with pytest.raises(KeyError):
        pkg.replace_part("word/ghost.xml", b"x")
    # create=True 允许新增，且追加在包尾
    pkg.replace_part("word/ghost.xml", b"x", create=True)
    assert pkg.list_parts()[-1] == "word/ghost.xml"


def test_read_part_unknown_raises(pkg):
    import pytest
    with pytest.raises(KeyError):
        pkg.read_part("word/ghost.xml")


def test_replace_text_in_part_count(pkg):
    n = pkg.replace_text_in_part("word/document.xml", "☑", "OK")
    assert n == 2
    assert b"\xe2\x98\x91" not in pkg.read_part("word/document.xml")
    assert pkg.read_part("word/document.xml").decode("utf-8").count("OK") == 2


def test_list_embeddings(pkg):
    assert pkg.list_embeddings() == ["word/embeddings/data.xlsx"]


def test_replace_embedded_with_file(pkg, tmp_path):
    new_xlsx = tmp_path / "new.xlsx"
    new_xlsx.write_bytes(b"FAKE_XLSX_BYTES")
    pkg.replace_embedded("word/embeddings/data.xlsx", str(new_xlsx))
    assert pkg.read_part("word/embeddings/data.xlsx") == b"FAKE_XLSX_BYTES"


def test_replace_embedded_with_bytes(pkg):
    pkg.replace_embedded("word/embeddings/data.xlsx", b"BYTES")
    assert pkg.read_part("word/embeddings/data.xlsx") == b"BYTES"


def test_replace_embedded_unknown_raises(pkg):
    import pytest
    with pytest.raises(KeyError):
        pkg.replace_embedded("word/embeddings/nothere.xlsx", b"x")


def test_load_from_file_object(docx_path):
    with open(docx_path, "rb") as f:
        pkg = DocxPackage(f)
    assert pkg.has_part("word/document.xml")
