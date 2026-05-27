# DOI 文献下载工具

批量 DOI → Zotero 导入 + Sci-Hub PDF 下载，Excel 管理全流程。

## 功能

1. 从 Excel 读取 DOI 列表
2. 通过 CrossRef API 获取文献元数据（标题、作者、期刊、年份等）
3. 通过 pyzotero 将文献导入 Zotero（可选）
4. 多源 PDF 下载：依次尝试 Unpaywall → PubMed Central → Semantic Scholar → Sci-Hub
5. 下载结果（成功/失败原因/下载来源）写回 Excel，支持断点续传

## 前置条件

- Python 3.10+
- Zotero 桌面版（如需导入功能）

## 安装

```bash
pip install -r requirements.txt
```

## 配置

### Zotero（可选，跳过则仅下载 PDF）

1. 访问 [zotero.org/settings/keys](https://www.zotero.org/settings/keys) 创建 API Key（需勾选写入权限）
2. 在同一页面找到你的 User ID
3. 打开 `doi_workflow.py`，修改 CONFIG 区域：

```python
ZOTERO_API_KEY = "你的_API_Key"
ZOTERO_USER_ID = "你的_User_ID"
ZOTERO_LIBRARY_TYPE = "user"
```

### 路径设置

```python
PDF_OUTPUT_DIR = r"./papers"      # PDF 保存目录
INPUT_EXCEL = r"./doi_list.xlsx"  # 输入 Excel 路径
```

## 用法

1. 准备 Excel 文件，确保有一列名为 `DOI`（也支持 `https://doi.org/10.xxx/xxx` 完整 URL 格式）
2. 运行脚本：

```bash
python doi_workflow.py
```

3. 等待处理完成，查看 Excel 中的结果

## Excel 输出说明

| 列名 | 说明 |
|------|------|
| DOI | 你输入的 DOI |
| 标题 | CrossRef 自动填充 |
| 期刊 | CrossRef 自动填充 |
| 年份 | CrossRef 自动填充 |
| PDF路径 | 下载成功后的本地路径 |
| 下载状态 | `已下载` / `下载失败` / `格式无效` |
| 备注 | 失败原因（如 Sci-Hub 未收录、网络超时等） |

## 断点续传

每处理完一条立即写入 Excel。如果脚本中断，重新运行会自动跳过已标记"已下载"的行。

## 多源下载策略

按以下优先级依次尝试，首个成功即停止：

| 优先级 | 来源 | 说明 | 配置要求 |
|--------|------|------|----------|
| 1 | **Unpaywall** | 全球最大的 OA 论文索引，覆盖各类 OA（金色、绿色、hybrid）期刊 | 需填写邮箱 |
| 2 | **PubMed Central** | 美国国立医学图书馆全文仓储，NIH 资助论文必须存缴 | 无需配置（可选填 API Key 提高限速） |
| 3 | **Semantic Scholar** | AI 学术搜索引擎，部分论文提供 OA PDF 链接 | 无需配置（可选填 API Key 提高限速） |
| 4 | **Sci-Hub** | 最后兜底，对老论文覆盖较好，新论文较少 | 无需配置 |

各源下载超时 30 秒，请求间隔 3 秒。

> **注意**：Google Scholar 不提供公开 API 且会主动反爬，因此未集成。Unpaywall + PMC 已覆盖绝大部分可免费获取的学术论文。

### 配置多源下载

```python
# 下载优先级（可调整顺序或移除某个源）
DOWNLOAD_SOURCES = ["unpaywall", "pmc", "semantic_scholar", "scihub"]

# Unpaywall 邮箱（必填，否则跳过此源）
UNPAYWALL_EMAIL = "your_email@example.com"

# 可选 API Key（提高速率限制）
NCBI_API_KEY = ""    # https://www.ncbi.nlm.nih.gov/account/
S2_API_KEY = ""      # https://www.semanticscholar.org/product/api
```

Excel 下载状态列会标注成功来源，如 `已下载(unpaywall)`、`已下载(pmc)` 等。

## 注意事项

- 较新的论文（2024 年后）Sci-Hub 可能尚未收录
- 请合理使用，遵守相关法律法规
- DOI 列中请勿混入 PubMed 等其他 URL

## License

MIT
