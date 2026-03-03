"""
公開說明書抽取 API (FastAPI)
=============================
接受上傳的 PDF 或 Google Drive 直連網址，
抽取「資金運用計畫」與「具體發債目的」並以 JSON 回傳。

部署於 Zeabur：
  - 監聽 PORT 環境變數（Zeabur 自動注入）
  - 支援 multipart/form-data 上傳 PDF
"""

import os
import re
import io
import tempfile
import httpx
import pdfplumber

from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.responses import JSONResponse
from contextlib import asynccontextmanager

# ──────────────────────────────────────────────────────────────
# 抽取邏輯（與 extract_prospectus.py 相同核心）
# ──────────────────────────────────────────────────────────────

TARGET_SECTIONS = [
    {
        "label": "資金運用計畫之用途及預計可能產生效益之概要",
        "keywords": [
            "資金運用計畫之用途",
            "資金運用計畫",
            "預計可能產生效益",
            "資金用途及預計",
        ],
        "detail_keywords": [
            "計畫項目及運用進度",
            "本計畫所需資金",
            "資金來源",
            "預計可能產生之效益",
        ],
    },
    {
        "label": "具體發債目的",
        "keywords": [
            "具體發債目的",
            "計畫項目及運用進度",
            "本次發行目的",
            "預計可能產生之效益",
        ],
        "detail_keywords": [
            "計畫項目及運用進度",
            "預計可能產生之效益",
        ],
    },
]

SECTION_END_PATTERNS = [
    r"^第\s*[一二三四五六七八九十百]+\s*[節章條款]",
    r"^[一二三四五六七八九十]+[、．.]\s*(?![\(（])\S",
    r"^\s*(?:附件|附錄|聲明|簽名|蓋章)",
]

SEE_PAGE_RE = re.compile(
    r"(?:請參閱|參閱|見|詳見|參見).*?第\s*(\d+)\s*頁", re.UNICODE
)


def extract_all_pages(pdf_bytes: bytes) -> list[dict]:
    result = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
            result.append({"page": i, "text": text, "lines": text.splitlines()})
    return result


def estimate_page_offset(pages: list[dict]) -> int:
    for pg in pages[:30]:
        lines = pg["lines"]
        if not lines:
            continue
        last_line = lines[-1].strip()
        if re.match(r"^\d+$", last_line):
            doc_page = int(last_line)
            return pg["page"] - doc_page
    return 0


def is_section_end(line: str, count: int) -> bool:
    if count < 4:
        return False
    stripped = line.strip()
    if not stripped:
        return False
    for pattern in SECTION_END_PATTERNS:
        if re.match(pattern, stripped):
            return True
    return False


def collect_from(pages, start_page_idx, start_line_idx, max_lines=200):
    collected = []
    blank_streak = 0
    for pg in pages[start_page_idx:]:
        lines = pg["lines"]
        line_start = start_line_idx if pg["page"] == pages[start_page_idx]["page"] else 0
        for ln in lines[line_start:]:
            if len(collected) >= max_lines:
                collected.append("...(內容過長，已截斷)")
                return "\n".join(collected).strip()
            if ln.strip() == "":
                blank_streak += 1
                if blank_streak > 4:
                    return "\n".join(collected).strip()
            else:
                blank_streak = 0
            if is_section_end(ln, len(collected)):
                return "\n".join(collected).strip()
            collected.append(ln)
    return "\n".join(collected).strip()


def search_in_pages(pages, keywords):
    for idx, pg in enumerate(pages):
        for li, line in enumerate(pg["lines"]):
            line_clean = line.strip().replace(" ", "")
            for kw in keywords:
                if kw.replace(" ", "") in line_clean:
                    return idx, li
    return None, None


def check_see_page_redirect(line: str):
    m = SEE_PAGE_RE.search(line)
    return int(m.group(1)) if m else None


def run_extraction(pdf_bytes: bytes) -> dict:
    pages = extract_all_pages(pdf_bytes)
    offset = estimate_page_offset(pages)
    results = {"total_pages": len(pages), "page_offset": offset, "sections": {}}

    for section_def in TARGET_SECTIONS:
        label = section_def["label"]
        keywords = section_def["keywords"]
        detail_keywords = section_def.get("detail_keywords", [])

        pg_idx, li_idx = search_in_pages(pages, keywords)

        if pg_idx is None:
            results["sections"][label] = {
                "found": False,
                "message": "未在文件中找到此區段",
            }
            continue

        found_page = pages[pg_idx]
        found_line_text = found_page["lines"][li_idx]

        # 偵測跳頁指示
        redirect_doc_page = None
        for scan_li in range(li_idx, min(li_idx + 6, len(found_page["lines"]))):
            redirect_doc_page = check_see_page_redirect(found_page["lines"][scan_li])
            if redirect_doc_page:
                break

        if redirect_doc_page:
            target_pdf_idx = redirect_doc_page + offset - 1
            if 0 <= target_pdf_idx < len(pages):
                detail_pg_idx, detail_li_idx = search_in_pages(
                    pages[target_pdf_idx: target_pdf_idx + 8],
                    detail_keywords if detail_keywords else keywords,
                )
                if detail_pg_idx is not None:
                    real_idx = target_pdf_idx + detail_pg_idx
                    content = collect_from(pages, real_idx, detail_li_idx)
                else:
                    content = collect_from(pages, target_pdf_idx, 0)
            else:
                content = collect_from(pages, pg_idx, li_idx)
        else:
            content = collect_from(pages, pg_idx, li_idx)

        results["sections"][label] = {
            "found": True,
            "first_hit_pdf_page": found_page["page"],
            "first_hit_line": found_line_text.strip(),
            "redirected_to_doc_page": redirect_doc_page,
            "content": content,
        }

    return results


# ──────────────────────────────────────────────────────────────
# FastAPI 應用程式
# ──────────────────────────────────────────────────────────────

app = FastAPI(
    title="公開說明書抽取 API",
    description="上傳 PDF 或提供 Google Drive 直連連結，自動抽取資金運用計畫與具體發債目的",
    version="1.0.0",
)


@app.get("/health")
def health_check():
    """N8N 可用此端點確認服務是否存活。"""
    return {"status": "ok"}


@app.post("/extract/upload")
async def extract_from_upload(file: UploadFile = File(...)):
    """
    上傳 PDF 檔案後抽取。
    N8N 使用「HTTP Request」節點，Body 選 Form Data，欄位名稱為 file。
    """
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="只接受 PDF 檔案")

    pdf_bytes = await file.read()
    if len(pdf_bytes) == 0:
        raise HTTPException(status_code=400, detail="檔案為空")

    try:
        result = run_extraction(pdf_bytes)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"抽取失敗：{str(e)}")

    return JSONResponse(content={"filename": file.filename, **result})


@app.post("/extract/from-url")
async def extract_from_url(url: str = Query(..., description="PDF 的直連網址（Google Drive export 格式）")):
    """
    從 URL 下載 PDF 後抽取。
    Google Drive 直連格式：
      https://drive.google.com/uc?export=download&id=<FILE_ID>

    N8N 使用「HTTP Request」節點，Method: POST，
    Query Parameter: url=<google_drive_export_url>
    """
    try:
        async with httpx.AsyncClient(follow_redirects=True, timeout=60) as client:
            response = await client.get(url)
            response.raise_for_status()
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"無法下載 PDF：{str(e)}")

    content_type = response.headers.get("content-type", "")
    if "pdf" not in content_type and not url.lower().endswith(".pdf") and "export" not in url:
        raise HTTPException(status_code=400, detail=f"回應內容不像 PDF（Content-Type: {content_type}）")

    pdf_bytes = response.content
    try:
        result = run_extraction(pdf_bytes)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"抽取失敗：{str(e)}")

    return JSONResponse(content={"source_url": url, **result})


# ──────────────────────────────────────────────────────────────
# 啟動（Zeabur 使用 PORT 環境變數）
# ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port)
