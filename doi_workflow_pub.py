"""
DOI 工作流：Excel 批量 DOI → Zotero 导入 + Sci-Hub PDF 下载

用法：
  1. 在下方 CONFIG 区域填入 Zotero API 凭证和路径
  2. 准备 Excel 文件，确保有一列名为 "DOI"
  3. 运行: python doi_workflow.py

前置条件：
  - 在 https://www.zotero.org/settings/keys 创建 API Key
  - 在 https://www.zotero.org/settings/keys 页面可看到 User ID
  - Zotero 桌面版已开启同步
"""

import os
import re
import sys
import time
import socket
import threading
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import urljoin

import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pyzotero import zotero
from tqdm import tqdm

# 全局 socket 超时，防止网络请求无限挂起
socket.setdefaulttimeout(60)

# ============================================================
# CONFIG — 按你的实际信息修改这里
# ============================================================

ZOTERO_API_KEY = "your_api_key_here"       # zotero.org/settings/keys
ZOTERO_USER_ID = "your_user_id_here"       # zotero.org/settings/keys 页面的 User ID
ZOTERO_LIBRARY_TYPE = "user"               # "user" 个人库 / "group" 群组库

PDF_OUTPUT_DIR = r"./papers"               # PDF 保存目录
INPUT_EXCEL = r"./doi_list.xlsx"           # 输入 Excel

DOI_COLUMN = "DOI"                         # Excel 中 DOI 列的列名

REQUEST_INTERVAL = 3                       # 请求间隔（秒），避免被限流
SCI_HUB_TIMEOUT = 30                       # Sci-Hub 单次请求超时（秒）
DOWNLOAD_TIMEOUT = 30                      # 通用 PDF 下载超时（秒）

# ============================================================
# 多源下载配置（按 DOWNLOAD_SOURCES 顺序依次尝试，成功即停）
# ============================================================

DOWNLOAD_SOURCES = ["unpaywall", "pmc", "semantic_scholar", "scihub"]

# Unpaywall（必填邮箱，否则自动跳过此源）
UNPAYWALL_EMAIL = "your_email@example.com"

# NCBI/PMC（可选，不填也能用但速率限制更低）
NCBI_API_KEY = ""

# Semantic Scholar（可选，不填也能用但速率限制更低）
S2_API_KEY = ""

# 来源显示名映射
SOURCE_DISPLAY = {
    "unpaywall": "Unpaywall",
    "pmc": "PMC",
    "semantic_scholar": "Semantic Scholar",
    "scihub": "Sci-Hub",
}

# ============================================================
# Sci-Hub 域名列表（按优先级排列，失效时可自行增减）
# ============================================================

SCI_HUB_DOMAINS = [
    "https://sci-hub.ru",
    "https://sci-hub.se",
    "https://sci-hub.st",
]


def sanitize_filename(name: str, max_len: int = 120) -> str:
    """清理文件名，移除非法字符并截断。"""
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    if len(name) > max_len:
        name = name[: max_len - 4].rstrip() + ".pdf"
    else:
        name = name + ".pdf"
    return name


def resolve_doi(doi: str, session: requests.Session) -> dict | None:
    """通过 CrossRef API 解析 DOI 获取文献元数据。"""
    url = f"https://api.crossref.org/works/{doi}"
    headers = {"Accept": "application/json"}
    try:
        resp = session.get(url, headers=headers, timeout=15)
        if resp.status_code == 404:
            return None
        resp.raise_for_status()
        data = resp.json()["message"]
    except Exception as e:
        print(f"  [!] CrossRef 查询失败: {e}")
        return None

    # 提取作者
    authors = []
    for a in data.get("author", [])[:10]:
        family = a.get("family", "")
        given = a.get("given", "")
        full = f"{family}, {given}" if family else given
        if full:
            authors.append(full)

    # 提取期刊名
    journal = ""
    for key in ("container-title", "short-container-title"):
        ct = data.get(key, [])
        if ct:
            journal = ct[0]
            break

    # 提取年份
    issued = data.get("issued", {})
    date_parts = issued.get("date-parts", [[None]])[0]
    year = date_parts[0] if date_parts else None

    return {
        "doi": data.get("DOI", doi),
        "title": data.get("title", [""])[0],
        "authors": authors,
        "journal": journal,
        "year": str(year) if year else "",
        "volume": data.get("volume", ""),
        "issue": data.get("issue", ""),
        "pages": data.get("page", ""),
        "issn": data.get("ISSN", [""])[0] if data.get("ISSN") else "",
        "url": data.get("URL", f"https://doi.org/{doi}"),
        "abstract": data.get("abstract", ""),
    }


def add_to_zotero(zot: zotero.Zotero, meta: dict) -> tuple[str | None, str]:
    """在 Zotero 中创建或查找期刊文章条目，返回 (item_key, action)。

    action 为: '创建' | '已存在' | '失败'
    """
    doi = meta["doi"]

    # 先查是否已存在同 DOI 条目
    try:
        existing = zot.items(q=doi, limit=1)
        if existing:
            key = existing[0]["key"]
            print(f"        Zotero 条目已存在 (key: {key})，跳过创建")
            return key, "已存在"
    except Exception:
        pass  # 查询失败则继续创建

    template = zot.item_template("journalArticle")
    template["title"] = meta["title"]
    template["DOI"] = doi
    template["url"] = meta["url"]
    template["publicationTitle"] = meta["journal"]
    template["volume"] = meta["volume"]
    template["issue"] = meta["issue"]
    template["pages"] = meta["pages"]
    template["ISSN"] = meta["issn"]
    template["abstractNote"] = meta["abstract"]
    template["date"] = meta["year"]

    # 作者
    creators = []
    for fullname in meta["authors"]:
        parts = fullname.split(", ", 1)
        if len(parts) == 2:
            creators.append({"creatorType": "author", "lastName": parts[0], "firstName": parts[1]})
        else:
            creators.append({"creatorType": "author", "name": fullname})
    template["creators"] = creators

    try:
        resp = zot.create_items([template])
        result = resp.get("success", {})
        if "0" in result:
            return result["0"], "创建"
        if isinstance(result, dict):
            for v in result.values():
                if isinstance(v, str) and len(v) == 8:
                    return v, "创建"
        return None, "失败"
    except Exception as e:
        print(f"  [!] Zotero 导入失败: {e}")
        return None, "失败"


def attach_pdf_to_zotero(zot: zotero.Zotero, parent_key: str, pdf_path: str) -> bool:
    """将 PDF 作为附件上传到 Zotero 条目。"""
    try:
        result = zot.attachment_simple([pdf_path], parent_key)
        return "success" in result
    except Exception as e:
        print(f"  [!] Zotero 附件上传失败: {e}")
        return False


def download_from_scihub(
    doi: str, output_dir: Path, filename: str, session: requests.Session
) -> tuple[str | None, str]:
    """
    尝试从 Sci-Hub 下载 PDF。

    返回: (pdf_path | None, status_message)
    """
    for domain in SCI_HUB_DOMAINS:
        for attempt in range(2):
            try:
                scihub_url = f"{domain}/{doi}"
                resp = session.get(
                    scihub_url,
                    timeout=SCI_HUB_TIMEOUT,
                    headers={
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36"
                    },
                )
                resp.raise_for_status()
                html = resp.text

                # 从页面中提取 PDF 直链
                pdf_url = _extract_pdf_url(html, domain)
                if not pdf_url:
                    continue

                # 下载 PDF
                pdf_resp = session.get(
                    pdf_url,
                    timeout=SCI_HUB_TIMEOUT,
                    headers={
                        "Referer": scihub_url,
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36",
                    },
                )
                content_type = pdf_resp.headers.get("Content-Type", "")

                # 验证是否为有效 PDF
                if "application/pdf" not in content_type:
                    with open(output_dir / (filename + ".error.html"), "wb") as f:
                        f.write(pdf_resp.content[:5000])
                    return None, f"Sci-Hub 返回非 PDF (Content-Type: {content_type})"

                if len(pdf_resp.content) < 10_000:
                    return None, "Sci-Hub 返回文件过小 (<10KB)，可能为错误页面"

                output_path = output_dir / filename
                with open(output_path, "wb") as f:
                    f.write(pdf_resp.content)

                return str(output_path), "已下载"

            except requests.Timeout:
                continue
            except requests.RequestException:
                continue

    return None, "下载失败：所有 Sci-Hub 域名均不可用"


def _extract_pdf_url(html: str, base_url: str) -> str | None:
    """从 Sci-Hub 页面提取 PDF 直链。"""
    # 方式 1: <meta name="citation_pdf_url" content="/storage/.../xxx.pdf">
    m = re.search(
        r'<meta\s+name=["\']citation_pdf_url["\']\s+content=["\']([^"\']+)["\']',
        html, re.IGNORECASE
    )
    if m:
        return urljoin(base_url, m.group(1))

    # 方式 2: <iframe src="...pdf">
    m = re.search(r'<iframe[^>]+src=["\']([^"\']+)["\']', html, re.IGNORECASE)
    if m:
        src = m.group(1)
        if "pdf" in src.lower() or "#view" in src.lower():
            return src

    # 方式 3: <embed src="...pdf">
    m = re.search(r'<embed[^>]+src=["\']([^"\']+)["\']', html, re.IGNORECASE)
    if m:
        return m.group(1)

    # 方式 4: <button onclick="location.href='...pdf'">
    m = re.search(
        r"location\.href\s*=\s*['\"]([^'\"]+\.pdf[^'\"]*)['\"]",
        html, re.IGNORECASE
    )
    if m:
        return m.group(1)

    # 方式 5: 查找含 pdf 的链接
    m = re.search(
        r'href=["\']((https?://[^"\']+?\.pdf[^"\']*?))["\']',
        html, re.IGNORECASE
    )
    if m:
        return m.group(1)

    return None


# ============================================================
# 多源下载：Unpaywall / PMC / Semantic Scholar
# ============================================================


def download_from_unpaywall(
    doi: str, output_dir: Path, filename: str, session: requests.Session
) -> tuple[str | None, str, str | None]:
    """通过 Unpaywall API 查找 OA PDF 并下载。

    返回: (pdf_path | None, status_message, landing_url | None)
    """
    if "your_" in UNPAYWALL_EMAIL or not UNPAYWALL_EMAIL:
        return None, "Unpaywall 未配置邮箱，跳过", None

    try:
        resp = session.get(
            f"https://api.unpaywall.org/v2/{doi}",
            params={"email": UNPAYWALL_EMAIL},
            timeout=15,
        )
        if resp.status_code == 404:
            return None, "Unpaywall 未收录此 DOI", None
        if resp.status_code == 403:
            return None, "Unpaywall 拒绝访问 (403)", None
        resp.raise_for_status()
        data = resp.json()
    except requests.Timeout:
        return None, "Unpaywall 请求超时", None
    except Exception as e:
        return None, f"Unpaywall 请求失败: {e}", None

    if not data.get("is_oa"):
        return None, "Unpaywall 无 OA 版本", None

    # 优先 best_oa_location，否则遍历 oa_locations 找 pdf
    pdf_url = None
    landing_url = None
    best = data.get("best_oa_location") or {}
    pdf_url = best.get("url_for_pdf")
    landing_url = best.get("url")

    if not pdf_url:
        for loc in data.get("oa_locations", []):
            pdf_url = loc.get("url_for_pdf")
            landing_url = landing_url or loc.get("url")
            if pdf_url:
                break

    if not pdf_url:
        msg = "Unpaywall 找到 OA 但无 PDF 直链"
        if landing_url:
            msg += f" (OA 页面: {landing_url})"
        return None, msg, landing_url

    # 下载 PDF
    try:
        pdf_resp = session.get(pdf_url, timeout=DOWNLOAD_TIMEOUT)
        content_type = pdf_resp.headers.get("Content-Type", "")
        if "application/pdf" not in content_type and not filename.endswith(".pdf"):
            return None, f"Unpaywall 返回非 PDF (Content-Type: {content_type})", landing_url
        if len(pdf_resp.content) < 10_000:
            return None, "Unpaywall 返回文件过小 (<10KB)", landing_url
        output_path = output_dir / filename
        with open(output_path, "wb") as f:
            f.write(pdf_resp.content)
        return str(output_path), "已下载(unpaywall)", None
    except requests.Timeout:
        return None, "Unpaywall PDF 下载超时", landing_url
    except Exception as e:
        return None, f"Unpaywall PDF 下载失败: {e}", landing_url


def download_from_pmc(
    doi: str, output_dir: Path, filename: str, session: requests.Session
) -> tuple[str | None, str]:
    """通过 NCBI E-utilities 从 PubMed Central 下载 PDF。"""
    # Step 1: DOI → PMCID
    idconv_params = {
        "ids": doi,
        "format": "json",
        "tool": "doi-workflow",
    }
    if NCBI_API_KEY:
        idconv_params["api_key"] = NCBI_API_KEY

    try:
        resp = session.get(
            "https://www.ncbi.nlm.nih.gov/pmc/utils/idconv/v1.0/",
            params=idconv_params,
            timeout=15,
        )
        resp.raise_for_status()
        data = resp.json()
    except requests.Timeout:
        return None, "PMC idconv 请求超时"
    except Exception as e:
        return None, f"PMC idconv 请求失败: {e}"

    records = data.get("records", [])
    if not records:
        return None, "PMC 未收录此 DOI"
    pmcid = records[0].get("pmcid")
    if not pmcid:
        return None, "PMC idconv 无 PMCID"

    # Step 2: 尝试直接下载 PDF
    efetch_params = {
        "db": "pmc",
        "id": pmcid,
        "rettype": "pdf",
        "tool": "doi-workflow",
    }
    if NCBI_API_KEY:
        efetch_params["api_key"] = NCBI_API_KEY

    try:
        pdf_resp = session.get(
            "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi",
            params=efetch_params,
            timeout=DOWNLOAD_TIMEOUT,
        )
        content_type = pdf_resp.headers.get("Content-Type", "")

        # 成功返回 PDF
        if "application/pdf" in content_type:
            if len(pdf_resp.content) < 10_000:
                return None, "PMC 返回文件过小 (<10KB)"
            output_path = output_dir / filename
            with open(output_path, "wb") as f:
                f.write(pdf_resp.content)
            return str(output_path), "已下载(pmc)"

        # 返回 XML 时尝试从中提取 PDF 链接
        if "xml" in content_type or pdf_resp.text.strip().startswith("<?xml"):
            pdf_url = _extract_pmc_xml_pdf(pdf_resp.text)
            if pdf_url:
                pdf2 = session.get(pdf_url, timeout=DOWNLOAD_TIMEOUT)
                ct2 = pdf2.headers.get("Content-Type", "")
                if "application/pdf" in ct2 and len(pdf2.content) >= 10_000:
                    output_path = output_dir / filename
                    with open(output_path, "wb") as f:
                        f.write(pdf2.content)
                    return str(output_path), "已下载(pmc)"

        return None, f"PMC 返回非 PDF (Content-Type: {content_type})"
    except requests.Timeout:
        return None, "PMC PDF 下载超时"
    except Exception as e:
        return None, f"PMC PDF 下载失败: {e}"


def _extract_pmc_xml_pdf(xml_text: str) -> str | None:
    """从 PMC XML 中提取 PDF 链接。"""
    try:
        root = ET.fromstring(xml_text)
        ns = {"xlink": "http://www.w3.org/1999/xlink"}
        for uri in root.iter("self-uri"):
            if uri.get("content-type") == "pdf":
                href = uri.get("{http://www.w3.org/1999/xlink}href")
                if href:
                    return href
        # 也尝试直接搜索 xlink:href="...pdf"
        m = re.search(r'xlink:href="(https?://[^"]+\.pdf)"', xml_text)
        if m:
            return m.group(1)
    except ET.ParseError:
        pass
    return None


def download_from_semantic_scholar(
    doi: str, output_dir: Path, filename: str, session: requests.Session
) -> tuple[str | None, str]:
    """通过 Semantic Scholar API 查找 OA PDF 并下载。"""
    headers = {}
    if S2_API_KEY:
        headers["x-api-key"] = S2_API_KEY

    try:
        resp = session.get(
            f"https://api.semanticscholar.org/graph/v1/paper/DOI:{doi}",
            params={"fields": "openAccessPdf"},
            headers=headers,
            timeout=15,
        )
        if resp.status_code == 404:
            return None, "Semantic Scholar 未收录此 DOI"
        if resp.status_code == 429:
            return None, "Semantic Scholar 速率限制 (429)"
        resp.raise_for_status()
        data = resp.json()
    except requests.Timeout:
        return None, "Semantic Scholar 请求超时"
    except Exception as e:
        return None, f"Semantic Scholar 请求失败: {e}"

    oa_pdf = data.get("openAccessPdf")
    if not oa_pdf or not oa_pdf.get("url"):
        return None, "Semantic Scholar 无 OA PDF"

    pdf_url = oa_pdf["url"]
    try:
        pdf_resp = session.get(pdf_url, timeout=DOWNLOAD_TIMEOUT)
        content_type = pdf_resp.headers.get("Content-Type", "")
        if "application/pdf" not in content_type and not filename.endswith(".pdf"):
            return None, f"Semantic Scholar 返回非 PDF (Content-Type: {content_type})"
        if len(pdf_resp.content) < 10_000:
            return None, "Semantic Scholar 返回文件过小 (<10KB)"
        output_path = output_dir / filename
        with open(output_path, "wb") as f:
            f.write(pdf_resp.content)
        return str(output_path), "已下载(semantic_scholar)"
    except requests.Timeout:
        return None, "Semantic Scholar PDF 下载超时"
    except Exception as e:
        return None, f"Semantic Scholar PDF 下载失败: {e}"


def _try_source(name: str, fn, doi: str, output_dir: Path, filename: str,
                session: requests.Session, cancel_event: threading.Event,
                result: dict):
    """在线程中尝试单个下载源。"""
    if cancel_event.is_set():
        return

    # 使用临时文件名避免多线程写入冲突
    tmp_filename = filename + f".tmp_{name}"
    try:
        ret = fn(doi, output_dir, tmp_filename, session)
    except Exception as e:
        ret = (None, f"{SOURCE_DISPLAY.get(name, name)} 异常: {e}")

    if len(ret) == 3:
        path, status, landing_url = ret
    else:
        path, status = ret
        landing_url = None

    if path and not cancel_event.is_set():
        # 首个成功者：重命名为最终文件名并广播取消
        final_path = output_dir / filename
        try:
            os.replace(path, final_path)
            path = str(final_path)
        except OSError:
            pass  # 重命名失败则保留临时文件名
        result["path"] = path
        result["status"] = status
        result["landing_url"] = None
        cancel_event.set()

    # 记录日志（按顺序输出，避免 tqdm 错乱）
    display = SOURCE_DISPLAY.get(name, name)
    if not cancel_event.is_set() or result.get("path") == path:
        pass  # 成功者在主线程输出

    result["errors"].append((name, status, landing_url))


def download_pdf_multi(
    doi: str, output_dir: Path, filename: str, session: requests.Session
) -> tuple[str | None, str]:
    """并行尝试所有下载源，首个成功即返回。"""
    sources = {
        "unpaywall": download_from_unpaywall,
        "pmc": download_from_pmc,
        "semantic_scholar": download_from_semantic_scholar,
        "scihub": download_from_scihub,
    }

    cancel_event = threading.Event()
    result = {"path": None, "status": None, "landing_url": None, "errors": []}

    executor = ThreadPoolExecutor(max_workers=len(DOWNLOAD_SOURCES))
    futures = []
    for source_name in DOWNLOAD_SOURCES:
        fn = sources.get(source_name)
        if not fn:
            continue
        futures.append(executor.submit(
            _try_source, source_name, fn, doi, output_dir, filename,
            session, cancel_event, result
        ))

    # 首个成功即停止等待，其余线程后台自行结束
    for f in as_completed(futures):
        f.result()
        if result["path"]:
            executor.shutdown(wait=False, cancel_futures=True)
            break
    else:
        executor.shutdown(wait=False)

    # 清理未使用的临时文件
    for name in DOWNLOAD_SOURCES:
        tmp_file = output_dir / (filename + f".tmp_{name}")
        if tmp_file.exists():
            try:
                tmp_file.unlink()
            except OSError:
                pass

    if result["path"]:
        return result["path"], result["status"]

    # 构建失败摘要，附带 OA 落地页链接
    errors = []
    fallback_urls = []
    for name, status, landing_url in result["errors"]:
        display = SOURCE_DISPLAY.get(name, name)
        errors.append(f"[{display}] {status}")
        if landing_url:
            fallback_urls.append(landing_url)

    summary = "；".join(errors) if errors else "所有源均未尝试"
    if fallback_urls:
        summary += f" | OA 落地页: {', '.join(fallback_urls[:3])}"
    return None, f"下载失败: {summary}"


def normalize_doi(doi: str) -> str:
    """从完整 URL 或裸字符串中提取裸 DOI。"""
    doi = doi.strip()
    # 匹配 https://doi.org/10.xxx/xxx 格式
    m = re.match(r"https?://(?:dx\.)?doi\.org/(10\.\d{4,}/.+)", doi, re.IGNORECASE)
    if m:
        return m.group(1)
    # 已是裸 DOI 则直接返回
    if re.match(r"^10\.\d{4,}/.+", doi):
        return doi
    return doi


def is_valid_doi(doi: str) -> bool:
    """基本 DOI 格式校验。"""
    doi = normalize_doi(doi)
    return bool(re.match(r"^10\.\d{4,}/.+", doi))


def get_zotero() -> zotero.Zotero | None:
    """创建并验证 Zotero 连接。未配置时返回 None。"""
    if "your_" in ZOTERO_API_KEY or "your_" in ZOTERO_USER_ID:
        print("[提示] Zotero API 未配置，将跳过 Zotero 导入，仅执行下载。")
        print("       如需导入 Zotero，请在脚本顶部 CONFIG 填入 API Key 和 User ID。")
        print("       获取方式: https://www.zotero.org/settings/keys\n")
        return None

    zot = zotero.Zotero(ZOTERO_USER_ID, ZOTERO_LIBRARY_TYPE, ZOTERO_API_KEY)
    try:
        zot.items(limit=1)
        print(f"[Zotero] 连接成功")
    except Exception as e:
        print(f"[警告] Zotero API 连接失败: {e}")
        print("        下载功能不受影响，但 Zotero 导入将被跳过。")
        return None
    return zot


def setup_excel(excel_path: str) -> tuple:
    """读取或创建 Excel，返回工作簿、工作表、DOI 列索引。"""
    if Path(excel_path).exists():
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "DOI列表"
        ws.append(["DOI", "标题", "期刊", "年份", "PDF路径", "下载状态", "备注"])
        # 设置表头样式
        for cell in ws[1]:
            cell.font = Font(bold=True)
        wb.save(excel_path)
        print(f"[创建] 已创建 Excel 模板: {excel_path}")
        print(f"       请在 'DOI' 列填入 DOI 号后重新运行脚本。")
        wb.close()
        sys.exit(0)

    # 找到 DOI 列
    doi_col_idx = None
    for col_idx, cell in enumerate(ws[1], start=1):
        if cell.value and DOI_COLUMN.lower() in str(cell.value).lower():
            doi_col_idx = col_idx
            break

    if doi_col_idx is None:
        # 在第一列之后插入 DOI 列
        ws.insert_cols(1)
        ws.cell(row=1, column=1, value=DOI_COLUMN)
        ws.cell(row=1, column=1).font = Font(bold=True)
        doi_col_idx = 1

    # 确保结果列存在
    result_headers = ["标题", "期刊", "年份", "PDF路径", "下载状态", "备注"]
    header_row = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    for i, h in enumerate(result_headers):
        if h not in header_row:
            ws.cell(row=1, column=ws.max_column + 1, value=h)
            ws.cell(row=1, column=ws.max_column).font = Font(bold=True)

    # 找到各列的索引
    header_cols = {}
    for col_idx in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col_idx).value
        if val:
            header_cols[str(val).strip()] = col_idx

    wb.save(excel_path)
    return wb, ws, doi_col_idx, header_cols


def main():
    print("=" * 60)
    print("  DOI 工作流：Excel → Zotero + 多源 PDF 下载")
    print("=" * 60)

    # 连接 Zotero
    zot = get_zotero()

    # 准备 PDF 输出目录
    output_dir = Path(PDF_OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    # 读取 Excel
    wb, ws, doi_col_idx, header_cols = setup_excel(INPUT_EXCEL)
    print(f"[Excel] 已加载: {INPUT_EXCEL}")

    # 收集所有 DOIs
    entries = []
    for row in range(2, ws.max_row + 1):
        doi_val = ws.cell(row=row, column=doi_col_idx).value
        if doi_val and str(doi_val).strip():
            doi = str(doi_val).strip()
            status_col = header_cols.get("下载状态", -1)
            current_status = ws.cell(row=row, column=status_col).value if status_col > 0 else None
            entries.append((row, doi, current_status))

    if not entries:
        print("Excel 中没有找到任何 DOI，请在 DOI 列填入 DOI 号后重试。")
        wb.close()
        return

    print(f"[统计] 共 {len(entries)} 条 DOI 待处理\n")

    # 创建共享 session（复用连接）
    session = requests.Session()

    # 逐条处理
    success_count = 0
    fail_count = 0
    skip_count = 0

    for row, doi, prev_status in tqdm(entries, desc="处理进度", unit="条"):
        print(f"\n{'-' * 50}")
        print(f"  DOI: {doi}")

        # 断点续传：跳过已下载的
        if prev_status and str(prev_status).startswith("已下载"):
            print(f"  [跳过] 已下载过")
            skip_count += 1
            continue

        # 规范化 DOI（从 URL 中提取裸 DOI）
        doi = normalize_doi(doi)

        # 校验 DOI 格式
        if not is_valid_doi(doi):
            _write_result(ws, row, header_cols, doi, status="格式无效", note="DOI 格式不正确")
            fail_count += 1
            wb.save(INPUT_EXCEL)
            continue

        # 1. CrossRef 解析元数据
        print(f"  [1/3] 查询 CrossRef 元数据...")
        meta = resolve_doi(doi, session)
        if meta is None:
            _write_result(ws, row, header_cols, doi, status="下载失败", note="CrossRef 元数据查询失败")
            fail_count += 1
            wb.save(INPUT_EXCEL)
            continue

        title = meta["title"]
        # 清理控制台下无法输出的字符（非断空格等）
        title = title.replace("\xa0", " ").replace("\xad", "")
        meta["title"] = title
        print(f"        标题: {title[:60]}...")

        # 写入元数据到 Excel
        _write_meta(ws, row, header_cols, meta)

        # 2. 导入 Zotero
        item_key = None
        if zot:
            print(f"  [2/3] 导入 Zotero...")
            item_key, zot_action = add_to_zotero(zot, meta)
            if item_key:
                if zot_action == "已存在":
                    print(f"        Zotero 条目已存在 (key: {item_key})，若下载成功将补充附件")
                else:
                    print(f"        Zotero 条目已创建 (key: {item_key})")
            else:
                print(f"        Zotero 创建失败，继续尝试下载 PDF")
        else:
            print(f"  [2/3] 跳过 Zotero (未配置)")

        # 3. 多源下载
        print(f"  [3/3] 多源下载 PDF...")
        if not title:
            safe_title = sanitize_filename(doi.replace("/", "_"))
        else:
            safe_title_part = f"{meta.get('year', '')}_{title}" if meta.get("year") else title
            safe_title = sanitize_filename(safe_title_part)

        pdf_path, status_msg = download_pdf_multi(doi, output_dir, safe_title, session)

        if pdf_path:
            # 尝试挂载到 Zotero
            if item_key and zot:
                attach_pdf_to_zotero(zot, item_key, pdf_path)

            _write_result(ws, row, header_cols, doi, status=status_msg, note="", pdf_path=pdf_path)
            print(f"        [OK] {status_msg} -> {pdf_path}")
            success_count += 1
        else:
            _write_result(ws, row, header_cols, doi, status="下载失败", note=status_msg)
            print(f"        [FAIL] {status_msg}")
            fail_count += 1

        # 即时保存
        wb.save(INPUT_EXCEL)

        # 请求间隔
        time.sleep(REQUEST_INTERVAL)

    wb.close()
    session.close()

    print(f"\n{'=' * 60}")
    print(f"  处理完成！")
    print(f"  成功下载: {success_count}  |  下载失败: {fail_count}  |  跳过: {skip_count}")
    print(f"  结果已写入: {INPUT_EXCEL}")
    print(f"  PDF 目录: {output_dir}")
    print(f"{'=' * 60}")


def _write_meta(ws, row: int, header_cols: dict, meta: dict):
    """写入文献元数据到 Excel。"""
    cols_map = {
        "标题": meta.get("title", ""),
        "期刊": meta.get("journal", ""),
        "年份": meta.get("year", ""),
    }
    for col_name, value in cols_map.items():
        col = header_cols.get(col_name)
        if col:
            ws.cell(row=row, column=col, value=value)
            ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True)


def _write_result(
    ws, row: int, header_cols: dict, doi: str,
    status: str, note: str, pdf_path: str | None = None
):
    """写入下载结果到 Excel。"""
    cols_map = {
        "DOI": doi,
        "下载状态": status,
        "备注": note,
    }
    if pdf_path:
        cols_map["PDF路径"] = pdf_path

    status_col = header_cols.get("下载状态")
    note_col = header_cols.get("备注")

    for col_name, value in cols_map.items():
        col = header_cols.get(col_name)
        if col:
            ws.cell(row=row, column=col, value=value)

    # 状态列着色
    if status_col:
        cell = ws.cell(row=row, column=status_col)
        if status.startswith("已下载"):
            cell.font = Font(color="006100")
            cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        elif "失败" in status or "无效" in status:
            cell.font = Font(color="9C0006")
            cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")


if __name__ == "__main__":
    main()
