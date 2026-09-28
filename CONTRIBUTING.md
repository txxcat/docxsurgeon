# 参与开发

这是个长期维护的小库，目标是**接口稳定、行为可预测、改动有测试兜底**。

## 环境

```bash
git clone https://github.com/txxcat/docxsurgeon
cd docxsurgeon
python -m venv .venv
# Windows: .venv\Scripts\activate     Linux/macOS: source .venv/bin/activate
pip install -e . pytest
```

要求 Python >= 3.9（代码里不要用 3.10+ 才有的语法，比如 `match`、`X | Y` 类型联合）。

## 跑测试

```bash
pytest -v
```

测试样本全部**动态生成**（`tests/conftest.py` 用 openpyxl + zipfile 现造一个最小 docx），
仓库里不放任何真实业务文档——这一点请务必保持。

## 代码约定

- **全内存**：任何新功能都不许落地临时目录、不许 `os.chdir`。这是本库区别于"解压改重打包"脚本的核心卖点。
- **保真**：`save()` 之后未修改的部件必须字节级不变。新增功能请补一条对应断言。
- **失败安全**：会修改包内容的上下文管理器，异常退出时不得写回半成品。
- **惰性导入**：重依赖（openpyxl）只在真正用到它那个方法里 import。
- 中文注释可以，但公开 API 的 docstring 要说清"为什么"而不只是"做什么"。

## 版本号

单一来源：只改 `src/docxsurgeon/package.py` 里的 `__version__`，
`pyproject.toml` 通过 `[tool.setuptools.dynamic]` 自动读取，不要在两处各写一遍。

## 发布到 PyPI

```bash
# 1) 改 __version__，并在 CHANGELOG 里加一段
# 2) 清理并构建
rm -rf dist build *.egg-info src/*.egg-info
python -m build
# 3) 先传测试源验证
python -m twine upload --repository testpypi dist/*
# 4) 正式发布
python -m twine upload dist/*
```

> 建议后续改成 **Trusted Publishing**（GitHub Actions + OIDC）：
> 不落盘 long-lived token，`on: release` 触发即可。当前先手动发。

## 提交规范

一次提交一件事，提交信息说清**动机和影响面**（改了行为？改了 API？修了什么 bug？）。
涉及行为变更的，务必在提交信息里写清"旧行为 → 新行为"。

## 已知边界（新功能别顺手"修"成另一个行为）

见 README 的「已知边界」一节。特别是：

- 图表缓存数据（`c:strCache` / `c:numCache`）**故意不同步改写**，这是待办而非疏漏。
- `open_embedded_xlsx` 依赖 openpyxl 往返，openpyxl 不建模的 xlsx 特性会丢——与直接用 openpyxl 编辑同一文件的风险一致，属于既有取舍。
