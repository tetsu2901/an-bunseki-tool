"""
案分析レポート生成ツール - Webアプリ

ブラウザからExcelファイルをアップロードすると、分析レポートを生成してダウンロードできる。
"""

import io
import os
import tempfile
from flask import Flask, request, send_file, render_template_string

from analyzer import generate_report

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB上限

HTML = '''
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>案分析レポート生成ツール</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background: #f0f4f8;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .container {
            background: white;
            border-radius: 16px;
            box-shadow: 0 4px 24px rgba(0,0,0,0.08);
            padding: 48px;
            max-width: 520px;
            width: 90%;
            text-align: center;
        }
        h1 {
            color: #1a365d;
            font-size: 24px;
            margin-bottom: 8px;
        }
        .subtitle {
            color: #718096;
            font-size: 14px;
            margin-bottom: 32px;
        }
        .drop-zone {
            border: 2px dashed #cbd5e0;
            border-radius: 12px;
            padding: 40px 24px;
            cursor: pointer;
            transition: all 0.2s;
            margin-bottom: 24px;
            position: relative;
        }
        .drop-zone:hover, .drop-zone.dragover {
            border-color: #4472C4;
            background: #ebf4ff;
        }
        .drop-zone .icon {
            font-size: 48px;
            margin-bottom: 12px;
        }
        .drop-zone p {
            color: #4a5568;
            font-size: 14px;
        }
        .drop-zone .filename {
            color: #2b6cb0;
            font-weight: bold;
            font-size: 15px;
            margin-top: 8px;
        }
        .drop-zone input[type="file"] {
            position: absolute;
            inset: 0;
            opacity: 0;
            cursor: pointer;
        }
        .btn {
            background: #4472C4;
            color: white;
            border: none;
            border-radius: 8px;
            padding: 14px 40px;
            font-size: 16px;
            font-weight: bold;
            cursor: pointer;
            transition: background 0.2s;
            width: 100%;
        }
        .btn:hover { background: #2b6cb0; }
        .btn:disabled {
            background: #a0aec0;
            cursor: not-allowed;
        }
        .status {
            margin-top: 16px;
            font-size: 14px;
            min-height: 20px;
        }
        .status.error { color: #e53e3e; }
        .status.success { color: #276749; }
        .status.processing { color: #4472C4; }
        .spinner {
            display: inline-block;
            width: 16px;
            height: 16px;
            border: 2px solid #4472C4;
            border-top-color: transparent;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
            vertical-align: middle;
            margin-right: 6px;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
    </style>
</head>
<body>
    <div class="container">
        <h1>案分析レポート生成ツール</h1>
        <p class="subtitle">時間枠一覧のExcelファイルをアップロードしてください</p>

        <form id="form" action="/upload" method="post" enctype="multipart/form-data">
            <div class="drop-zone" id="dropZone">
                <div class="icon">📊</div>
                <p>ここにファイルをドラッグ&ドロップ<br>またはクリックして選択</p>
                <div class="filename" id="fileName"></div>
                <input type="file" name="file" id="fileInput" accept=".xlsx,.xlsm">
            </div>
            <button type="submit" class="btn" id="submitBtn" disabled>レポート生成</button>
        </form>

        <div class="status" id="status"></div>
    </div>

    <script>
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const fileName = document.getElementById('fileName');
        const submitBtn = document.getElementById('submitBtn');
        const status = document.getElementById('status');
        const form = document.getElementById('form');

        fileInput.addEventListener('change', () => {
            if (fileInput.files.length > 0) {
                fileName.textContent = fileInput.files[0].name;
                submitBtn.disabled = false;
                status.textContent = '';
                status.className = 'status';
            }
        });

        ['dragover', 'dragenter'].forEach(evt => {
            dropZone.addEventListener(evt, e => {
                e.preventDefault();
                dropZone.classList.add('dragover');
            });
        });
        ['dragleave', 'drop'].forEach(evt => {
            dropZone.addEventListener(evt, () => {
                dropZone.classList.remove('dragover');
            });
        });

        form.addEventListener('submit', async (e) => {
            e.preventDefault();
            if (!fileInput.files.length) return;

            submitBtn.disabled = true;
            status.className = 'status processing';
            status.innerHTML = '<span class="spinner"></span>処理中...';

            const formData = new FormData(form);
            try {
                const res = await fetch('/upload', { method: 'POST', body: formData });
                if (res.ok) {
                    const blob = await res.blob();
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    const cd = res.headers.get('Content-Disposition') || '';
                    const match = cd.match(/filename\*=UTF-8''(.+)/);
                    a.download = match ? decodeURIComponent(match[1]) : 'レポート.xlsx';
                    a.href = url;
                    a.click();
                    URL.revokeObjectURL(url);
                    status.className = 'status success';
                    status.textContent = 'レポートのダウンロードが開始されました';
                } else {
                    const text = await res.text();
                    status.className = 'status error';
                    status.textContent = 'エラー: ' + text;
                }
            } catch (err) {
                status.className = 'status error';
                status.textContent = 'エラー: ' + err.message;
            }
            submitBtn.disabled = false;
        });
    </script>
</body>
</html>
'''


@app.route('/')
def index():
    return render_template_string(HTML)


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return 'ファイルが送信されていません', 400

    file = request.files['file']
    if file.filename == '':
        return 'ファイルが選択されていません', 400

    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ('.xlsx', '.xlsm'):
        return 'xlsx または xlsm ファイルをアップロードしてください', 400

    # 一時ファイルに保存して処理
    tmpdir = tempfile.mkdtemp()
    try:
        input_path = os.path.join(tmpdir, file.filename)
        file.save(input_path)

        base = os.path.splitext(file.filename)[0]
        output_name = f'{base}_分析レポート.xlsx'
        output_path = os.path.join(tmpdir, output_name)

        generate_report(input_path, output_path)

        # ファイルをメモリに読み込んでから一時ファイルを削除
        with open(output_path, 'rb') as f:
            data = io.BytesIO(f.read())

        return send_file(
            data,
            as_attachment=True,
            download_name=output_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
    except Exception as e:
        return f'処理エラー: {e}', 500
    finally:
        import shutil
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == '__main__':
    print('=' * 50)
    print('  案分析レポート生成ツール')
    print('  http://localhost:5050')
    print('=' * 50)
    app.run(host='0.0.0.0', port=5050, debug=False)
