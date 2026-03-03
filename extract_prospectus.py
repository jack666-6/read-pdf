"""
公開說明書 PDF 資訊抽取工具
==============================
抽取以下區段：
  1. 本次資金運用計畫之用途及預計可能產生效益之概要
     （含自動追蹤「請參閱第 XX 頁」之跳頁指示）
  2. 具體發債目的
     （若無獨立標題，自動從「本次發行計畫」中萃取）

使用方式：
    python extract_prospectus.py <PDF路徑>
    例如：python extract_prospectus.py 公開說明書.pdf

依賴套件（需先安裝）：
    pip install pdfplumber --break-system-packages
"""

import sys
import re
import pdfplumber


# ──────────────────────────────────────────────────────────────
# 搜尋設定
# ──────────────────────────────────────────────────────────────

# 目標區段：從這些關鍵字找到區段起始位置
TARGET_SECTIONS = [
    {
        "label": "【本次資金運用計畫之用途及預計可能產生效益之概要】",
        "keywords": [
            "資金運用計畫之用途",
            "資金運用計畫",
            "預計可能產生效益",
            "資金用途及預計",
        ],
        # 區段真正展開的標題關鍵字（有些文件是「摘要頁→詳細頁」）
        "detail_keywords": [
            "計畫項目及運用進度",
            "本計畫所需資金",
            "資金來源",
            "預計可能產生之效益",
        ],
    },
    {
        "label": "【具體發債目的（本次發行計畫項目及預計效益）】",
        "keywords": [
            "具體發債目的",
            "計畫項目及運用進度",   # 通常包含發債目的的詳細說明
            "本次發行目的",
            "預計可能產生之效益",   # 效益段通常緊接在發債目的之後
        ],
        "detail_keywords": [
            "計畫項目及運用進度",
            "預計可能產生之效益",
        ],
    },
]

# 區段結束的信號（遇到這些就停止收集）
# 注意：只有「主要章節」才算結束，小節序號（3. 4.）不算
SECTION_END_PATTERNS = [
    r"^第\s*[一二三四五六七八九十百]+\s*[節章條款]",          # 第X節/章
    r"^[一二三四五六七八九十]+[、．.]\s*(?![\(（])\S",          # 一、二、（但不是 一、(一)）
    r"^\s*(?:附件|附錄|聲明|簽名|蓋章)",                        # 附件/附錄 等結尾區塊
]

# 「參閱第 XX 頁」正則（抓取跳頁指示）
SEE_PAGE_RE = re.compile(
    r"(?:請參閱|參閱|見|詳見|參見).*?第\s*(\d+)\s*頁", re.UNICODE
)

# 文件頁碼與 PDF 頁碼的偏移（通常公開說明書封面不計頁碼）
# 程式會自動估算，也可在此手動設定（0=不偏移）
PAGE_OFFSET_HINT = None   # None = 自動估算


# ──────────────────────────────────────────────────────────────
# 工具函數
# ──────────────────────────────────────────────────────────────

def extract_all_pages(pdf_path: str) -> list[dict]:
    """讀取全部頁面，回傳 [{page(1-based), text, lines}, ...]。"""
    result = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total = len(pdf.pages)
            print(f"📄 共 {total} 頁，開始讀取...\n")
            for i, page in enumerate(pdf.pages, start=1):
                text = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
                result.append({
                    "page": i,
                    "text": text,
                    "lines": text.splitlines(),
                })
    except FileNotFoundError:
        print(f"❌ 找不到檔案：{pdf_path}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ 讀取 PDF 時發生錯誤：{e}")
        sys.exit(1)
    return result


def estimate_page_offset(pages: list[dict]) -> int:
    """
    估算文件內頁碼與 PDF 實際頁碼的差距。
    原理：找到第一個頁腳有頁碼數字的頁面，算出偏移量。
    """
    for pg in pages[:30]:   # 只看前 30 頁
        lines = pg["lines"]
        if not lines:
            continue
        last_line = lines[-1].strip()
        if re.match(r"^\d+$", last_line):
            doc_page = int(last_line)
            offset = pg["page"] - doc_page
            return offset
    return 0


def doc_page_to_pdf_index(doc_page: int, offset: int) -> int:
    """將文件頁碼轉為 pages[] 的 0-based 索引。"""
    return doc_page + offset - 1


def is_section_end(line: str, content_count: int) -> bool:
    """判斷是否進入下一個主要章節（區段結束）。"""
    if content_count < 4:
        return False
    stripped = line.strip()
    if not stripped:
        return False
    for pattern in SECTION_END_PATTERNS:
        if re.match(pattern, stripped):
            return True
    return False


def collect_from(pages: list[dict], start_page_idx: int,
                 start_line_idx: int, max_lines: int = 150) -> str:
    """
    從指定頁/行開始往後收集文字，直到遇到下一主節或行數上限。
    """
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


def search_in_pages(pages: list[dict], keywords: list[str]):
    """在所有頁面中搜尋關鍵字，回傳第一個匹配的 (page_index, line_index)。"""
    for idx, pg in enumerate(pages):
        for li, line in enumerate(pg["lines"]):
            line_clean = line.strip().replace(" ", "")
            for kw in keywords:
                if kw.replace(" ", "") in line_clean:
                    return idx, li
    return None, None


def check_see_page_redirect(line: str) -> int | None:
    """若該行有「請參閱第 XX 頁」，回傳頁碼，否則回傳 None。"""
    m = SEE_PAGE_RE.search(line)
    if m:
        return int(m.group(1))
    return None


# ──────────────────────────────────────────────────────────────
# 主流程
# ──────────────────────────────────────────────────────────────

def extract_sections(pdf_path: str):
    print("=" * 60)
    print("  公開說明書資訊抽取程式")
    print("=" * 60)
    print(f"檔案：{pdf_path}\n")

    pages = extract_all_pages(pdf_path)

    # 估算頁碼偏移
    if PAGE_OFFSET_HINT is None:
        offset = estimate_page_offset(pages)
    else:
        offset = PAGE_OFFSET_HINT
    print(f"📌 偵測到頁碼偏移量：{offset}（文件頁碼 = PDF頁碼 - {offset}）\n")

    results = {}

    for section_def in TARGET_SECTIONS:
        label = section_def["label"]
        keywords = section_def["keywords"]
        detail_keywords = section_def.get("detail_keywords", [])

        print(f"🔍 搜尋區段：{label}")

        pg_idx, li_idx = search_in_pages(pages, keywords)

        if pg_idx is None:
            print(f"   ⚠️  未找到此區段（嘗試關鍵字：{keywords}）\n")
            results[label] = None
            continue

        found_page = pages[pg_idx]
        found_line = found_page["lines"][li_idx]
        print(f"   ✅ 首次命中於 PDF 第 {found_page['page']} 頁：{found_line.strip()[:60]}")

        # ── 檢查是否是「請參閱第 XX 頁」的摘要行 ──────────────────
        # 向後掃描最多 5 行尋找跳頁指示
        redirect_doc_page = None
        for scan_li in range(li_idx, min(li_idx + 6, len(found_page["lines"]))):
            redirect_doc_page = check_see_page_redirect(found_page["lines"][scan_li])
            if redirect_doc_page:
                break

        if redirect_doc_page:
            target_pdf_idx = doc_page_to_pdf_index(redirect_doc_page, offset)
            print(f"   🔄 偵測到跳頁指示 → 文件第 {redirect_doc_page} 頁（PDF索引 {target_pdf_idx + 1}）")

            if 0 <= target_pdf_idx < len(pages):
                # 在目標頁面上找到詳細內容的起點
                detail_pg_idx, detail_li_idx = search_in_pages(
                    pages[target_pdf_idx: target_pdf_idx + 8],   # 只在目標頁附近搜
                    detail_keywords if detail_keywords else keywords
                )
                if detail_pg_idx is not None:
                    real_idx = target_pdf_idx + detail_pg_idx
                    print(f"   📋 詳細內容起始於 PDF 第 {pages[real_idx]['page']} 頁")
                    content = collect_from(pages, real_idx, detail_li_idx, max_lines=200)
                else:
                    # 直接從目標頁開始收集
                    content = collect_from(pages, target_pdf_idx, 0, max_lines=200)
            else:
                print(f"   ⚠️  跳頁目標超出範圍，改從命中行開始收集")
                content = collect_from(pages, pg_idx, li_idx)
        else:
            # 無跳頁指示，從命中行直接收集
            content = collect_from(pages, pg_idx, li_idx, max_lines=200)

        results[label] = {
            "page": found_page["page"],
            "redirect": redirect_doc_page,
            "content": content,
        }
        print()

    # ── 輸出結果 ─────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("  抽取結果")
    print("=" * 60)

    for label, result in results.items():
        print(f"\n{label}")
        print("-" * 55)
        if result is None:
            print("（未找到此區段，請確認文件中是否有此章節）")
        else:
            loc_note = f"PDF 第 {result['page']} 頁"
            if result["redirect"]:
                loc_note += f"（摘要）→ 詳細內容於文件第 {result['redirect']} 頁"
            print(f"位置：{loc_note}\n")
            print(result["content"])
        print()

    return results


# ──────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python extract_prospectus.py <PDF路徑>")
        print("範例：python extract_prospectus.py 公開說明書.pdf")
        sys.exit(1)

    extract_sections(sys.argv[1])
