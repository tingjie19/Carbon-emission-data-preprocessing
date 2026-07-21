"""
總分類帳分割腳本

從總分類帳目錄找出水電瓦斯費、燃料費所在頁數，
並從總分類帳中擷取對應頁面另存，最後一頁蓋上大小章。

用法：
    python3 process.py <資料夾路徑> [民國年] [大小章資料夾路徑]

例：
    python3 process.py "/Users/tim/Desktop/碳健檢/總分類帳/傳燈" 114
    python3 process.py "/Users/tim/Desktop/碳健檢/總分類帳/傳燈" 114 "/Users/tim/Desktop/碳健檢/總分類帳/大小章"

資料夾內需包含：
    {公司名}{民國年}總分類帳目錄.pdf
    {公司名}{民國年}總分類帳.pdf

大小章資料夾（預設與總分類帳同層的「大小章」資料夾）需包含：
    去背_大章.png
    去背_小章.png

輸出（存至同資料夾）：
    {公司名}{民國年}水電瓦斯.pdf  （若目錄有水電瓦斯費）
    {公司名}{民國年}汽油.pdf      （若目錄有燃料費）
"""

import sys
import os
import re
import io
import tempfile
import pdfplumber
from pypdf import PdfReader, PdfWriter

# ── 關鍵字設定 ──────────────────────────────────────────────────────────────
# 關鍵字（用於找出候選會計項目） → 輸出檔名後綴 → 需 100% 完全吻合的正式科目名稱
# 會計項目名稱須與 exact_names 其中之一完全相同才會併入主檔；
# 僅包含關鍵字但夾雜其他字元（例如「營業成本-水電瓦斯費」）視為不同科目，
# 另存為獨立 PDF，不與主檔混在一起。
TARGETS = [
    (['水電瓦斯'],       '水電瓦斯', ['水電瓦斯費']),
    (['燃料', '汽油'],   '汽油',     ['燃料費', '汽油費']),
]


def extract_catalog(catalog_path):
    """
    解析目錄 PDF，回傳排序後的 [(名稱, 頁碼), ...] 列表。
    頁碼為總分類帳的 1-based 頁碼。
    """
    with pdfplumber.open(catalog_path) as pdf:
        full_text = '\n'.join(p.extract_text() or '' for p in pdf.pages)

    # 抓取「名稱 數字」格式
    entries = re.findall(r'([^\s\d（][^\d\n]{0,15}?)\s+(\d+)', full_text)
    seen = {}
    for name, page in entries:
        name = name.strip()
        p = int(page)
        if name not in seen:
            seen[name] = p

    return sorted(seen.items(), key=lambda x: x[1])


def get_page_ranges(keywords, entries_sorted):
    """
    從排序後的目錄條目找出所有含 keywords 之候選會計項目的頁碼範圍
    （同一關鍵字可能對應多個科目名稱，例如「水電瓦斯費」與
    「營業成本-水電瓦斯費」皆含「水電瓦斯」，須全部找出，
    是否併入主檔則由呼叫端依 exact_names 100% 比對決定）。
    結尾頁 = 下一個條目的起始頁 - 1。
    回傳 [(name, start, end), ...]，找不到則回傳空列表。
    """
    ranges = []
    for i, (name, start) in enumerate(entries_sorted):
        if any(kw in name for kw in keywords):
            end = entries_sorted[i + 1][1] - 1 if i + 1 < len(entries_sorted) else start
            ranges.append((name, start, end))
    return ranges


def apply_stamps(output_path, stamp_dir):
    """
    在 PDF 最後一頁右下角蓋大小章。
    左大章（4.5×4.5 cm）、右小章（3×3 cm），距頁緣 1.5 cm，兩章間距 0.5 cm。
    stamp_dir 需包含 去背_大章.png 和 去背_小章.png。
    """
    try:
        from PIL import Image
        from reportlab.pdfgen import canvas as rl_canvas
        from reportlab.lib.units import cm
    except ImportError:
        print("  （略過蓋章：需安裝 Pillow 和 reportlab）")
        return

    da_path   = os.path.join(stamp_dir, '去背_大章.png')
    xiao_path = os.path.join(stamp_dir, '去背_小章.png')
    if not os.path.exists(da_path) or not os.path.exists(xiao_path):
        print(f"  （略過蓋章：找不到大小章圖片於 {stamp_dir}）")
        return

    DA_SIZE   = 4.5 * cm
    XIAO_SIZE = 3.0 * cm
    MARGIN    = 1.5 * cm
    GAP       = 0.5 * cm

    reader    = PdfReader(output_path)
    last_page = reader.pages[-1]
    page_w    = float(last_page.mediabox.width)
    page_h    = float(last_page.mediabox.height)

    # 左大章、右小章，靠右下角
    xiao_x = page_w - MARGIN - XIAO_SIZE
    xiao_y = MARGIN
    da_x   = xiao_x - GAP - DA_SIZE
    da_y   = MARGIN

    packet = io.BytesIO()
    c = rl_canvas.Canvas(packet, pagesize=(page_w, page_h))

    def draw_stamp(img_path, x, y, size):
        img = Image.open(img_path).convert('RGBA')
        tmp = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
        img.save(tmp.name)
        tmp.close()
        c.drawImage(tmp.name, x, y, width=size, height=size,
                    preserveAspectRatio=True, mask='auto')
        os.unlink(tmp.name)

    draw_stamp(da_path,   da_x,   da_y,   DA_SIZE)
    draw_stamp(xiao_path, xiao_x, xiao_y, XIAO_SIZE)
    c.save()
    packet.seek(0)

    stamp_reader = PdfReader(packet)
    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        if i == len(reader.pages) - 1:
            page.merge_page(stamp_reader.pages[0])
        writer.add_page(page)

    with open(output_path, 'wb') as f:
        writer.write(f)


def extract_pages(ledger_path, ranges, output_path, stamp_dir=None):
    """從總分類帳擷取指定頁範圍列表（1-based，可多段）合併另存，最後一頁蓋章。"""
    reader = PdfReader(ledger_path)
    writer = PdfWriter()
    total  = len(reader.pages)

    for start_page, end_page in ranges:
        end_page = min(end_page, total)
        for i in range(start_page - 1, end_page):
            writer.add_page(reader.pages[i])

    with open(output_path, 'wb') as f:
        writer.write(f)

    if stamp_dir:
        apply_stamps(output_path, stamp_dir)

    pages_desc = '、'.join(f'第 {s}～{e} 頁' for s, e in ranges)
    print(f"  → {os.path.basename(output_path)}（{pages_desc}，含蓋章）")


def detect_company_and_year(folder_path, roc_year):
    """從資料夾名稱推測公司名；民國年由參數傳入。"""
    folder_name = os.path.basename(os.path.abspath(folder_path))
    return folder_name, roc_year


def find_pdf(folder, company, year, suffix):
    """在 folder 中找符合 '{company}{year}{suffix}.pdf' 的檔案。
    對「總分類帳目錄」另外接受「總分類帳封面」作為別名。
    """
    # 精確比對
    name = f"{company}{year}{suffix}.pdf"
    path = os.path.join(folder, name)
    if os.path.exists(path):
        return path
    # 容錯：掃描資料夾找含 suffix 的 PDF
    for f in os.listdir(folder):
        if suffix in f and f.endswith('.pdf'):
            return os.path.join(folder, f)
    # 若找目錄，也接受「封面」
    if '目錄' in suffix:
        alias = suffix.replace('目錄', '封面')
        for f in os.listdir(folder):
            if alias in f and f.endswith('.pdf'):
                return os.path.join(folder, f)
    return None


def find_stamp_dir(folder_path):
    """自動尋找大小章資料夾，往上最多三層尋找。"""
    path = os.path.abspath(folder_path)
    for _ in range(3):
        path = os.path.dirname(path)
        candidate = os.path.join(path, '大小章')
        if os.path.isdir(candidate):
            return candidate
    return None


def scan_ledger_by_category(ledger_path):
    """
    無目錄時，直接掃描總分類帳每頁右上角的「會計項目」，
    回傳排序後的 [(名稱, 頁碼), ...] 列表（1-based）。
    每頁文字第3行（index 2）即為會計項目。
    """
    entries = []
    with pdfplumber.open(ledger_path) as pdf:
        prev_cat = None
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ''
            lines = [l.strip() for l in text.split('\n') if l.strip()]
            cat = lines[2] if len(lines) > 2 else ''
            if cat and cat != prev_cat:
                entries.append((cat, i + 1))  # 1-based
                prev_cat = cat
    return entries


def run(folder_path, roc_year=114, stamp_dir=None, company_name=None):
    if company_name:
        company, year = company_name, roc_year
    else:
        company, year = detect_company_and_year(folder_path, roc_year)
    print(f"公司: {company}，民國年: {year}")

    catalog_path = find_pdf(folder_path, company, year, '總分類帳目錄')
    ledger_path  = find_pdf(folder_path, company, year, '總分類帳')

    if not ledger_path:
        print("找不到總分類帳 PDF（需含「總分類帳」但不含「目錄」）")
        sys.exit(1)
    if '目錄' in os.path.basename(ledger_path) or '封面' in os.path.basename(ledger_path):
        print("找到的總分類帳 PDF 含「目錄」或「封面」，請確認資料夾中有完整的總分類帳 PDF")
        sys.exit(1)

    # 自動尋找大小章資料夾
    if stamp_dir is None:
        stamp_dir = find_stamp_dir(folder_path)
    if stamp_dir:
        print(f"大小章: {stamp_dir}")
    else:
        print("找不到大小章資料夾，輸出 PDF 將不蓋章")

    if catalog_path:
        print(f"目錄: {os.path.basename(catalog_path)}")
        print(f"總分類帳: {os.path.basename(ledger_path)}")
        print()
        entries = extract_catalog(catalog_path)
    else:
        print(f"（無目錄 PDF，改從總分類帳各頁右上角會計項目掃描）")
        print(f"總分類帳: {os.path.basename(ledger_path)}")
        print()
        entries = scan_ledger_by_category(ledger_path)

    found_any = False
    for keywords, output_suffix, exact_names in TARGETS:
        matches = get_page_ranges(keywords, entries)
        if not matches:
            print(f"  未找到含 {'、'.join(keywords)} 的條目，略過")
            continue

        main_ranges = [(s, e) for name, s, e in matches if name in exact_names]
        extra_by_name: dict[str, list] = {}
        for name, s, e in matches:
            if name not in exact_names:
                extra_by_name.setdefault(name, []).append((s, e))

        if main_ranges:
            out_name = f"{company}{year}{output_suffix}.pdf"
            out_path = os.path.join(folder_path, out_name)
            extract_pages(ledger_path, main_ranges, out_path, stamp_dir)
            found_any = True
        else:
            print(f"  找不到與「{'、'.join(exact_names)}」完全吻合的科目，主檔略過")

        for name, name_ranges in extra_by_name.items():
            print(f"  科目「{name}」含關鍵字但非完全吻合，另存獨立 PDF")
            out_name = f"{company}{year}{name}.pdf"
            out_path = os.path.join(folder_path, out_name)
            extract_pages(ledger_path, name_ranges, out_path, stamp_dir)
            found_any = True

    if found_any:
        print("\n完成！")
    else:
        print("\n未找到水電瓦斯費或燃料費，請確認總分類帳格式。")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    folder_path = sys.argv[1]
    roc_year    = int(sys.argv[2]) if len(sys.argv) > 2 else 114
    stamp_dir   = sys.argv[3] if len(sys.argv) > 3 else None
    run(folder_path, roc_year, stamp_dir)
