# docxsurgeon

[![CI](https://github.com/txxcat/docxsurgeon/actions/workflows/ci.yml/badge.svg)](https://github.com/txxcat/docxsurgeon/actions/workflows/ci.yml)
[![License](https://img.shields.io/github/license/txxcat/docxsurgeon)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.9%20%7C%203.10%20%7C%203.11%20%7C%203.12%20%7C%203.13-blue)](https://www.python.org)
[![Status](https://img.shields.io/badge/status-alpha-orange)](#roadmap)

**Operate on the *inside* of .docx files — replace embedded Excel, update chart ranges, fix checkbox glyphs. All in memory.**

**对 docx 做 zip 包级"手术"：替换/编辑嵌入的 xlsx、更新图表数据引用范围、修复复选框打印符号。全程内存操作，不落地临时目录。**

## 为什么需要它

OOXML 文件（docx/xlsx/pptx）本质是 zip 包，里面藏着 `word/embeddings/*.xlsx`（嵌入的 Excel）、`word/charts/chart*.xml`（图表定义）、`xl/tables/*.xml`（Excel 表格范围）。这些**包内结构**在主流库里无人认领：

| 需求 | python-docx | docxtpl | openpyxl | docxsurgeon |
|---|---|---|---|---|
| 替换 docx 内嵌入的 xlsx | ❌ | ⚠️ 需走 Jinja2 render 流程 | ❌（只认独立文件） | ✅ |
| **编辑**嵌入 xlsx 的单元格 | ❌ | ❌ | ❌ | ✅ 内存直改，退出自动写回 |
| 改嵌入 xlsx 里 Excel 表格范围 | ❌ | ❌ | ❌ | ✅ |
| 更新图表 XML 数据引用行数 | ❌ | ❌ | ❌ | ✅ |
| Unicode ☑ → Wingdings 符号 run（打印修复） | ❌ | ❌ | ❌ | ✅ |
| 文档模板变量填充 | ❌ | ✅ | — | ❌（不抢 docxtpl 的活） |
| 文档模型读写（段落/表格） | ✅ | ✅ | — | ❌（不抢 python-docx 的活） |

一句话定位：**python-docx 管文档模型，docxtpl 管模板填充，docxsurgeon 管"已有文档的包级修改"。**

核心逻辑提取自一套在广电播控机房生产环境运行多年的报表工具，经过真实 Word / WPS 文档的长期验证。

## 安装

要求 **Python >= 3.9**。

**从 PyPI（待发布）**

> 本项目尚未在 PyPI 正式发布，0.1.0 上线后会启用下面的命令。当前请使用下方的源码安装方式。

```bash
pip install docxsurgeon
```

**从源码安装（当前阶段推荐）**

```bash
pip install "git+https://github.com/txxcat/docxsurgeon.git"
```

**开发模式**

```bash
git clone https://github.com/txxcat/docxsurgeon
cd docxsurgeon
pip install -e .
```

依赖说明：运行时仅依赖 `openpyxl>=3.0`（编辑嵌入 xlsx 时用到，已随安装自动拉取）。其余功能（部件读写、图表范围改写、符号修复）不依赖任何第三方库。

## 快速上手

```python
from docxsurgeon import DocxPackage, set_table_range

pkg = DocxPackage("周报模板.docx")

# 1) 直接编辑包内嵌入的 Excel（内存中，无临时目录）
with pkg.open_embedded_xlsx("word/embeddings/Microsoft_Excel____1.xlsx") as wb:
    ws = wb.active
    ws["A2"] = "9月"
    ws["B2"] = 12
    set_table_range(ws, "A1:C8")   # 同步 Excel 表格范围

# 2) 图表引用行数随数据条数伸缩（Sheet1!$B$2:$B$20 -> $B$2:$B$8）
pkg.set_chart_data_rows("word/charts/chart2.xml", last_row=8)

# 3) 替换嵌入文件本体
pkg.replace_embedded("word/embeddings/Microsoft_Word_Document3.docx", "新附件.docx")

# 4) 修复 ☑ 打印乱码：替换为 Wingdings 2 符号 run
pkg.fix_checkbox_glyph()

pkg.save("输出.docx")   # 未修改的部件字节级原样保留
```

## 进阶示例

**探索文档结构** —— 动刀前先看看包里有什么：

```python
pkg = DocxPackage("周报模板.docx")
print(pkg.list_parts())        # 全部部件，按原始 zip 顺序
print(pkg.list_embeddings())   # 只列 word/embeddings/ 下的嵌入文件
print(pkg.has_part("word/document.xml"))   # True
```

**部件内文本替换** —— 把某个 XML 部件当纯文本做子串替换，返回替换次数：

```python
n = pkg.replace_text_in_part(
    "word/charts/chart2.xml",
    "Sheet1!$B$2:$B$20",
    "Sheet1!$B$2:$B$8",
)
print(f"改了 {n} 处")
```

**失败安全** —— `open_embedded_xlsx` 的 with 块里若抛异常，本次对嵌入簿的修改会被整体丢弃，包保持原样：

```python
try:
    with pkg.open_embedded_xlsx("word/embeddings/data.xlsx") as wb:
        wb.active["A1"] = "新值"
        raise RuntimeError("模拟中途出错")
except RuntimeError:
    pass
# 此时嵌入簿未被破坏，可安全地改做别的处理或直接 save()
```

## API

### `DocxPackage(source)`

`source` 为文件路径或二进制文件对象。

| 方法 | 说明 |
|---|---|
| `list_parts()` / `has_part(name)` / `read_part(name)` | zip 部件级读取 |
| `replace_part(name, data, create=False)` | 整体替换部件（bytes/str） |
| `replace_text_in_part(name, old, new)` | 部件内文本替换，返回次数 |
| `list_embeddings()` | 列出 `word/embeddings/` 下的全部嵌入文件 |
| `replace_embedded(name, source)` | 用本地文件或字节替换嵌入文件 |
| `open_embedded_xlsx(name)` | **上下文管理器**：yield openpyxl Workbook，正常退出自动写回，异常退出放弃修改 |
| `set_chart_data_rows(chart_part, last_row, sheet="Sheet1", start_row=2)` | 改写图表数据引用范围，返回改写数 |
| `fix_checkbox_glyph(part, search="☑", symbol_font="Wingdings 2", symbol_char="0052", ...)` | Unicode 勾选符 → 符号 run |
| `save(path)` | 写回磁盘 |

### 模块函数

- `set_table_range(ws, ref)` —— 在 `open_embedded_xlsx` 块内修改 Excel 表格（ListObject）范围。
- `__version__` —— 当前版本号（与 `package.py` 单一来源一致）。

## 设计要点

- **全内存**：部件读写走 `BytesIO`，没有临时目录、没有 `os.chdir`、没有竞态清理。
- **保真**：`save()` 按原 zip 顺序写入；未修改的部件字节级不变（测试保证）。
- **失败安全**：`open_embedded_xlsx` 中抛异常则放弃本次修改，包保持原样。
- **惰性依赖**：只有用到 `open_embedded_xlsx` 才需要 openpyxl。

## 已知边界

- 图表的缓存数据（`c:strCache` / `c:numCache`）不会同步改写——Word 打开时会从嵌入工作簿刷新。若你的场景要求"不刷新也立即显示新值"，参考 roadmap。
- 嵌入的 PDF/OLE 对象（`oleObjectNNN.bin`）不支持替换（OLE 是编码格式）。
- `open_embedded_xlsx` 依赖 openpyxl 往返，openpyxl 不建模的 xlsx 特性会丢失（与直接用 openpyxl 编辑同一文件的风险一致）。

## Roadmap

- [ ] 上传 PyPI 正式版（当前未发布，`pip install` 需走源码）
- [ ] 图表缓存数据（strCache/numCache）双写
- [ ] 嵌入 xlsx 的图表部件（`word/charts` 引用嵌入簿的完整链路更新）
- [ ] pptx / xlsx 宿主包的一等支持（当前 API 以 docx 命名，机制通用）
- [ ] CLI：`python -m docxsurgeon list/extract/replace ...`

## 开发 & 贡献

```bash
git clone https://github.com/txxcat/docxsurgeon
cd docxsurgeon
pip install -e . pytest
pytest
```

仓库通过 GitHub Actions 跑 CI（见 `.github/workflows/ci.yml`，每次 push/PR 自动跑测试）。

想参与贡献？请先阅读 [CONTRIBUTING.md](CONTRIBUTING.md)。

## License

[MIT](LICENSE)
