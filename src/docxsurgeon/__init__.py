"""docxsurgeon：对 OOXML 文件（docx/xlsx）做 zip 包级"手术"。"""
from .package import DocxPackage, set_table_range, __version__

__all__ = ["DocxPackage", "set_table_range", "__version__"]
