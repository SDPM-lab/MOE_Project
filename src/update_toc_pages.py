"""依排版後的 PDF 產生目次頁碼（src/toc_pages.json），並回報貳至肆章頁數。

用法：
    node src/build_docx.js                       # 產生 docx 與 src/toc_entries.json
    soffice --headless --convert-to pdf <docx>   # 以 LibreOffice 轉 PDF（或用 Word 另存 PDF）
    python3 src/update_toc_pages.py <pdf>        # 寫入 src/toc_pages.json
    node src/build_docx.js                       # 以正確頁碼重新產生 docx
"""
import json
import re
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
TOC_START_PAGE = 3  # 封面與目次之後才開始搜尋標題


def page_texts(pdf):
    out = subprocess.run(["pdftotext", "-layout", str(pdf), "-"], capture_output=True, text=True, check=True).stdout
    return [re.sub(r"\s+", "", p) for p in out.split("\f")]


def main(pdf):
    entries = json.loads((HERE / "toc_entries.json").read_text(encoding="utf-8"))
    pages = page_texts(pdf)
    result, cursor = {}, TOC_START_PAGE - 1
    for e in entries:
        key = re.sub(r"\s+", "", e["text"])[:14]
        for i in range(cursor, len(pages)):
            if key in pages[i]:
                result[e["text"]] = i + 1
                cursor = i
                break
        else:
            sys.exit(f"heading not found: {e['text']}")
    (HERE / "toc_pages.json").write_text(json.dumps(result, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    start = result["貳、全程計畫摘要"]
    end = result["伍、116年度計畫經費需求"] - 1
    print(f"total pages: {len(pages) - (1 if pages and not pages[-1] else 0)}")
    print(f"貳至肆: pages {start}-{end} = {end - start + 1} pages")


if __name__ == "__main__":
    main(sys.argv[1])
