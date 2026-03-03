"""
公開說明書抽取 API (FastAPI)
"""

import os
import re
import tempfile
import httpx
import pdfplumber

from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.responses import JSONResponse

TARGET_SECTIONS = [
    {
        "label": "資金運用計畫之用途及預計可能產生效益之概要",
        "keywords": ["資金運用計畫之用途","資金運用計畫","預計可能產生效益","資金用途及預計"],
        "detail_keywords": ["計畫項目及運用進度","本計畫所需資金","資金來源","預計可能產生之效益"],
    },
]

SECTION_END_PATTERNS = [
    r"^第\s*[一二三四五六七八九十百]+\s*[節章條款]",
    r"^[一二三四五六七八九十]+[、．.]\s*(?![\(（])\S",
    r"^\s*(?:附件|附錄|聲明|簽名|蓋章)",
]

SEE_PAGE_RE = re.compile(r"(?:請參閱|參閱|見|詳見|參見).*?第\s*(\d+)\s*頁", re.UNICODE)


def extract_all_pages(pdf_path):
    result = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
            result.append({"page": i, "text": text, "lines": text.splitlines()})
            page.flush_cache()
    return result

def estimate_page_offset(pages):
    for pg in pages[:30]:
        lines = pg["lines"]
        if not lines:
            continue
        last_line = lines[-1].strip()
        if re.match(r"^\d+$", last_line):
            return pg["page"] - int(last_line)
    return 0

def is_section_end(line, count):
    if count < 4 or not line.strip():
        return False
    for pattern in SECTION_END_PATTERNS:
        if re.match(pattern, line.strip()):
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

def check_see_page_redirect(line):
    m = SEE_PAGE_RE.search(line)
    return int(m.group(1)) if m else None

def parse_summary(content):
    summary = {"發行資金總額": None, "發行用途": []}

    # 抓總額
    total_match = re.search(r"本計畫所需資金總額[：:]\s*新台幣\s*([\d,，]+)\s*仟元", content)
    if total_match:
        summary["發行資金總額"] = f"新台幣 {total_match.group(1).replace('，', ',')} 仟元"

    # 抓計畫項目區塊
    plan_match = re.search(r"計畫項目及運用進度.*?\n(.*?)合計", content, re.DOTALL)
    if plan_match:
        block = plan_match.group(1)
        lines = block.splitlines()
        skip_words = ["合計", "小計", "總額", "單位", "計畫項目", "預定完成", "季度", "進度"]
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            # 純中文行（項目名稱）接下一行有金額
            if re.match(r"^[\u4e00-\u9fff\-－()（）\s]{2,20}$", line) and len(line) >= 2:
                if not any(w in line for w in skip_words):
                    if i + 1 < len(lines):
                        amt = re.search(r"([\d,]{4,})", lines[i + 1])
                        if amt:
                            summary["發行用途"].append(f"{line}：{amt.group(1).replace(',','')} 仟元")
            # 中文+金額同一行
            elif re.search(r"[\u4e00-\u9fff]{2,}", line):
                m = re.search(r"([\u4e00-\u9fff\s]{2,15})\s+\d{3,}.*?\s+([\d,]{4,})\s+[\d,]+", line)
                if m and not any(w in m.group(1) for w in skip_words):
                    summary["發行用途"].append(f"{m.group(1).strip()}：{m.group(2).replace(',','')} 仟元")
            i += 1

    return summary

def run_extraction(pdf_path):
    pages = extract_all_pages(pdf_path)
    offset = estimate_page_offset(pages)
    sd = TARGET_SECTIONS[0]
    pg_idx, li_idx = search_in_pages(pages, sd["keywords"])
    if pg_idx is None:
        return {"發行資金總額": None, "發行用途": [], "error": "未找到資金運用計畫區段"}

    found_page = pages[pg_idx]
    redirect_doc_page = None
    for scan_li in range(li_idx, min(li_idx + 6, len(found_page["lines"]))):
        redirect_doc_page = check_see_page_redirect(found_page["lines"][scan_li])
        if redirect_doc_page:
            break

    if redirect_doc_page:
        target_pdf_idx = redirect_doc_page + offset - 1
        if 0 <= target_pdf_idx < len(pages):
            dpg, dli = search_in_pages(pages[target_pdf_idx:target_pdf_idx+8], sd["detail_keywords"])
            if dpg is not None:
                content = collect_from(pages, target_pdf_idx + dpg, dli)
            else:
                content = collect_from(pages, target_pdf_idx, 0)
        else:
            content = collect_from(pages, pg_idx, li_idx)
    else:
        content = collect_from(pages, pg_idx, li_idx)

    return parse_summary(content)


app = FastAPI(title="公開說明書抽取 API", version="2.0.0")

@app.get("/health")
def health_check():
    return {"status": "ok"}

@app.post("/extract/upload")
async def extract_from_upload(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="只接受 PDF 檔案")
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp_path = tmp.name
            total = 0
            while True:
                chunk = await file.read(1024 * 1024)
                if not chunk:
                    break
                tmp.write(chunk)
                total += len(chunk)
        if total == 0:
            raise HTTPException(status_code=400, detail="檔案為空")
        result = run_extraction(tmp_path)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"抽取失敗：{str(e)}")
    finally:
        if tmp_path:
            try: os.unlink(tmp_path)
            except: pass
    return JSONResponse(content={"filename": file.filename, **result})

@app.post("/extract/from-url")
async def extract_from_url(url: str = Query(...)):
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp_path = tmp.name
            async with httpx.AsyncClient(follow_redirects=True, timeout=120) as client:
                async with client.stream("GET", url) as response:
                    response.raise_for_status()
                    async for chunk in response.aiter_bytes(chunk_size=1024 * 1024):
                        tmp.write(chunk)
        result = run_extraction(tmp_path)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"抽取失敗：{str(e)}")
    finally:
        if tmp_path:
            try: os.unlink(tmp_path)
            except: pass
    return JSONResponse(content={"source_url": url, **result})

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port)
