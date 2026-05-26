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


def add_to_zotero(zot: zotero.Zotero, meta: dict) -> str | None:
    """在 Zotero 中创建期刊文章条目，返回 item key。"""
    template = zot.item_template("journalArticle")
    template["title"] = meta["title"]
    template["DOI"] = meta["doi"]
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
        # pyzotero 返回格式: {"success": {"0": "item_key"}} 或 {"0": "item_key"}
        if "0" in result:
            return result["0"]
        if isinstance(result, dict):
            for v in result.values():
                if isinstance(v, str) and len(v) == 8:
                    return v
        return None
    except Exception as e:
        print(f"  [!] Zotero 导入失败: {e}")
        return None


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
    print("  DOI 工作流：Excel → Zotero + Sci-Hub 下载")
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
        if prev_status and "已下载" in str(prev_status):
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
            item_key = add_to_zotero(zot, meta)
            if item_key:
                print(f"        Zotero 条目已创建 (key: {item_key})")
            else:
                print(f"        Zotero 创建失败，继续尝试下载 PDF")
        else:
            print(f"  [2/3] 跳过 Zotero (未配置)")

        # 3. Sci-Hub 下载
        print(f"  [3/3] Sci-Hub 下载 PDF...")
        if not title:
            safe_title = sanitize_filename(doi.replace("/", "_"))
        else:
            safe_title_part = f"{meta.get('year', '')}_{title}" if meta.get("year") else title
            safe_title = sanitize_filename(safe_title_part)

        pdf_path, status_msg = download_from_scihub(doi, output_dir, safe_title, session)

        if pdf_path:
            # 尝试挂载到 Zotero
            if item_key and zot:
                attach_pdf_to_zotero(zot, item_key, pdf_path)

            _write_result(ws, row, header_cols, doi, status="已下载", note="", pdf_path=pdf_path)
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
        if status == "已下载":
            cell.font = Font(color="006100")
            cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        elif "失败" in status or "无效" in status:
            cell.font = Font(color="9C0006")
            cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")


if __name__ == "__main__":
    main()
