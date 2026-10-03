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

发布走 **Trusted Publishing**（GitHub Actions + OIDC），仓库里**不存任何长期 token**。
工作流：`.github/workflows/publish.yml`（`on: release: published` 触发）。

**一次性前置配置**（项目首次发布前完成）：

1. PyPI 账号开启 **2FA**（PyPI 要求）。
2. 在 <https://pypi.org/manage/account/publishing/> 添加 **pending publisher**：
   - PyPI project name: `docxsurgeon`
   - Owner: `txxcat`
   - Repository name: `docxsurgeon`
   - Workflow name: `publish.yml`
   - Environment name: `pypi`
3. 在本仓库 Settings → Environments 新建名为 `pypi` 的 environment。

> pending publisher 只在**第一次真正上传时**才创建项目、并自动转成正式 publisher；它**不预留名字**，
> 若发布前被别人抢注了同名就会失效——所以确认名字空了尽早发。

**发版流程**：

```bash
# 1) 改 src/docxsurgeon/package.py 里的 __version__（版本号单一来源）
# 2) 本地验证
pytest -v && python -m build
# 3) 提交并打 tag
git commit -am "release: v0.1.0"
git tag v0.1.0
git push origin main --tags
# 4) 在 GitHub 上基于该 tag 建 Release 并 Publish —— 这一步才触发上传
```

> PyPI 不允许重传同一版本号。发错了只能在 PyPI 上 yank，然后改版本号重新发布。

## 提交规范

一次提交一件事，提交信息说清**动机和影响面**（改了行为？改了 API？修了什么 bug？）。
涉及行为变更的，务必在提交信息里写清"旧行为 → 新行为"。

## 已知边界（新功能别顺手"修"成另一个行为）

见 README 的「已知边界」一节。特别是：

- 图表缓存数据（`c:strCache` / `c:numCache`）**故意不同步改写**，这是待办而非疏漏。
- `open_embedded_xlsx` 依赖 openpyxl 往返，openpyxl 不建模的 xlsx 特性会丢——与直接用 openpyxl 编辑同一文件的风险一致，属于既有取舍。
