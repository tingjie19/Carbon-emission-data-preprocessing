"""
碳健檢資料前處理 Web App

使用：
    pip install -r requirements.txt
    python app.py
然後瀏覽器會自動開啟 http://localhost:5000
"""

import io
import os
import sys
import uuid
import zipfile
import tempfile
import threading
import webbrowser
import contextlib

from flask import Flask, jsonify, render_template, request, send_file

# ── 路徑設定（相容 PyInstaller 打包後的 frozen 模式）──────────────────────────
if getattr(sys, 'frozen', False):
    # 打包後：資源在 sys._MEIPASS 暫存目錄
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SCRIPTS_DIR  = os.path.join(BASE_DIR, 'scripts')
TEMPLATE_DIR = os.path.join(BASE_DIR, '範例')

# 讓 Flask 找到 templates 資料夾
app = Flask(__name__, template_folder=os.path.join(BASE_DIR, 'templates'))

# 加入 scripts 到模組搜尋路徑，以便直接 import
if SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, SCRIPTS_DIR)

from split_ledger      import run as split_run   # noqa: E402
from process_utilities import run as util_run    # noqa: E402
from process_gasoline  import run as gas_run     # noqa: E402

# 暫存處理結果（本地單人工具，記憶體存即可）
_results: dict[str, bytes] = {}


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    roc_year     = (request.form.get('year') or '114').strip()
    ad_year      = int(roc_year) + 1911
    ledger_files = request.files.getlist('ledger_folder')
    stamp_files  = request.files.getlist('stamp_folder')

    if not ledger_files or all(f.filename == '' for f in ledger_files):
        return jsonify(success=False, log='請選擇總分類帳資料夾。')

    log_lines = []

    with tempfile.TemporaryDirectory() as tmpdir:
        # ── 還原所有上傳檔案（平坦結構） ───────────────────────────────
        work_dir = os.path.join(tmpdir, 'work')
        os.makedirs(work_dir, exist_ok=True)

        for f in ledger_files:
            fname = f.filename.replace('\\', '/').split('/')[-1]
            if fname:
                f.save(os.path.join(work_dir, fname))

        # ── 還原大小章資料夾 ────────────────────────────────────────────
        stamp_dir_path = None
        has_stamps = False
        for f in stamp_files:
            fname = f.filename.replace('\\', '/').split('/')[-1]
            if fname:
                if stamp_dir_path is None:
                    stamp_dir_path = os.path.join(tmpdir, '大小章')
                    os.makedirs(stamp_dir_path, exist_ok=True)
                f.save(os.path.join(stamp_dir_path, fname))
                has_stamps = True

        # ── 找出所有總分類帳 PDF（命名規則：{客戶名稱}{年}總分類帳.pdf） ──
        ledger_pdfs = sorted([
            f for f in os.listdir(work_dir)
            if f.endswith('.pdf') and '總分類帳' in f
               and '目錄' not in f and '封面' not in f
        ])

        if not ledger_pdfs:
            return jsonify(
                success=False,
                log='找不到總分類帳 PDF（需含「總分類帳」但不含「目錄」或「封面」）。'
            )

        company_outputs: dict[str, list[str]] = {}

        for ledger_fname in ledger_pdfs:
            # 從檔名提取客戶名稱：移除尾綴 {year}總分類帳.pdf
            suffix = f'{roc_year}總分類帳.pdf'
            company = (ledger_fname[:-len(suffix)]
                       if ledger_fname.endswith(suffix)
                       else os.path.splitext(ledger_fname)[0])

            before = set(os.listdir(work_dir))

            # ── 步驟一：分割 + 蓋章 ────────────────────────────────────
            log_lines.append(f'\n=== [{company}] 步驟一：分割總分類帳 + 蓋章 ===')
            buf = io.StringIO()
            try:
                with contextlib.redirect_stdout(buf):
                    split_run(work_dir, int(roc_year),
                              stamp_dir_path if has_stamps else None,
                              company_name=company)
            except SystemExit:
                pass
            except Exception as e:
                buf.write(f'[錯誤] {e}\n')
            log_lines.append(buf.getvalue())

            # ── 步驟二：水電瓦斯 Excel ──────────────────────────────────
            util_pdf = os.path.join(work_dir, f'{company}{roc_year}水電瓦斯.pdf')
            if os.path.exists(util_pdf):
                log_lines.append(f'\n=== [{company}] 步驟二：水電瓦斯 Excel ===')
                buf = io.StringIO()
                try:
                    with contextlib.redirect_stdout(buf):
                        util_run(util_pdf, TEMPLATE_DIR, work_dir,
                                 int(roc_year), ad_year, company_name=company)
                except Exception as e:
                    buf.write(f'[錯誤] {e}\n')
                log_lines.append(buf.getvalue())

            # ── 步驟三：燃料 Excel ──────────────────────────────────────
            gas_pdf = os.path.join(work_dir, f'{company}{roc_year}汽油.pdf')
            if os.path.exists(gas_pdf):
                log_lines.append(f'\n=== [{company}] 步驟三：燃料 Excel ===')
                buf = io.StringIO()
                try:
                    with contextlib.redirect_stdout(buf):
                        gas_run(gas_pdf, TEMPLATE_DIR, work_dir,
                                int(roc_year), ad_year, company_name=company)
                except Exception as e:
                    buf.write(f'[錯誤] {e}\n')
                log_lines.append(buf.getvalue())

            # 記錄此客戶產出的新檔案
            after = set(os.listdir(work_dir))
            company_outputs[company] = [
                f for f in (after - before)
                if f.endswith('.xlsx') or f.endswith('.pdf')
            ]

        # ── 打包：每家公司放進自己的子資料夾 ──────────────────────────
        output_files = [f for files in company_outputs.values() for f in files]

        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for company, files in company_outputs.items():
                for fname in files:
                    zf.write(os.path.join(work_dir, fname), f'{company}/{fname}')
        zbuf.seek(0)

        job_id = str(uuid.uuid4())
        _results[job_id] = zbuf.read()

        log_lines.append(f'\n完成！共處理 {len(ledger_pdfs)} 位客戶，產出 {len(output_files)} 個檔案。')

    return jsonify(
        success=True,
        log='\n'.join(log_lines),
        job_id=job_id,
        file_count=len(output_files)
    )


@app.route('/download/<job_id>')
def download(job_id):
    data = _results.pop(job_id, None)
    if data is None:
        return '找不到結果，請重新執行。', 404
    return send_file(
        io.BytesIO(data),
        as_attachment=True,
        download_name='碳健檢結果.zip',
        mimetype='application/zip'
    )


def _open_browser():
    """延遲 1.5 秒後自動開啟瀏覽器（讓 Flask 先啟動）"""
    import time
    time.sleep(1.5)
    webbrowser.open('http://localhost:7860')


if __name__ == '__main__':
    print('碳健檢前處理工具啟動中...')
    print('瀏覽器將自動開啟，或手動前往 http://localhost:5000')
    threading.Thread(target=_open_browser, daemon=True).start()
    app.run(debug=False, port=7860)
