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
        # ── 還原總分類帳資料夾 ──────────────────────────────────────────
        raw_names    = [f.filename.replace('\\', '/') for f in ledger_files if f.filename]
        company_name = raw_names[0].split('/')[0] if raw_names else 'company'
        company_dir  = os.path.join(tmpdir, company_name)
        os.makedirs(company_dir, exist_ok=True)

        for f in ledger_files:
            fname = f.filename.replace('\\', '/').split('/')[-1]
            if fname:
                f.save(os.path.join(company_dir, fname))

        # ── 還原大小章資料夾 ────────────────────────────────────────────
        stamp_dir = os.path.join(tmpdir, '大小章')
        os.makedirs(stamp_dir, exist_ok=True)
        has_stamps = False
        for f in stamp_files:
            fname = f.filename.replace('\\', '/').split('/')[-1]
            if fname:
                f.save(os.path.join(stamp_dir, fname))
                has_stamps = True

        # ── 步驟一：分割 + 蓋章 ────────────────────────────────────────
        log_lines.append('=== 步驟一：分割總分類帳 + 蓋章 ===')
        buf = io.StringIO()
        try:
            with contextlib.redirect_stdout(buf):
                split_run(company_dir, int(roc_year),
                          stamp_dir if has_stamps else None)
        except SystemExit:
            pass  # split_ledger 在找不到檔案時會 sys.exit，此處攔截
        except Exception as e:
            buf.write(f'[錯誤] {e}\n')
        log_lines.append(buf.getvalue())

        # ── 步驟二：水電瓦斯 Excel ──────────────────────────────────────
        util_pdf = next(
            (os.path.join(company_dir, f)
             for f in os.listdir(company_dir)
             if '水電瓦斯' in f and f.endswith('.pdf')),
            None
        )
        if util_pdf:
            log_lines.append('\n=== 步驟二：水電瓦斯 Excel ===')
            buf = io.StringIO()
            try:
                with contextlib.redirect_stdout(buf):
                    util_run(util_pdf, TEMPLATE_DIR, company_dir,
                             int(roc_year), ad_year)
            except Exception as e:
                buf.write(f'[錯誤] {e}\n')
            log_lines.append(buf.getvalue())

        # ── 步驟三：燃料 Excel ──────────────────────────────────────────
        gas_pdf = next(
            (os.path.join(company_dir, f)
             for f in os.listdir(company_dir)
             if '汽油' in f and f.endswith('.pdf')),
            None
        )
        if gas_pdf:
            log_lines.append('\n=== 步驟三：燃料 Excel ===')
            buf = io.StringIO()
            try:
                with contextlib.redirect_stdout(buf):
                    gas_run(gas_pdf, TEMPLATE_DIR, company_dir,
                            int(roc_year), ad_year)
            except Exception as e:
                buf.write(f'[錯誤] {e}\n')
            log_lines.append(buf.getvalue())

        # ── 打包結果 ────────────────────────────────────────────────────
        output_files = [
            f for f in os.listdir(company_dir)
            if f.endswith('.xlsx') or
               (f.endswith('.pdf') and ('水電瓦斯' in f or '汽油' in f))
        ]

        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, 'w', zipfile.ZIP_DEFLATED) as zf:
            for fname in output_files:
                zf.write(os.path.join(company_dir, fname), fname)
        zbuf.seek(0)

        job_id = str(uuid.uuid4())
        _results[job_id] = zbuf.read()

        log_lines.append(f'\n完成！共 {len(output_files)} 個輸出檔案。')

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
    webbrowser.open('http://localhost:5000')


if __name__ == '__main__':
    print('碳健檢前處理工具啟動中...')
    print('瀏覽器將自動開啟，或手動前往 http://localhost:5000')
    threading.Thread(target=_open_browser, daemon=True).start()
    app.run(debug=False, port=5000)
