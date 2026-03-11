"""
案分析レポート生成ツール - GUIアプリ

Excelファイルをドラッグ&ドロップまたはファイル選択で読み込み、
分析レポートを自動生成する。
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import threading

from analyzer import generate_report


class App:
    def __init__(self, root):
        self.root = root
        self.root.title('案分析レポート生成ツール')
        self.root.geometry('520x360')
        self.root.resizable(False, False)
        self.root.configure(bg='#f0f4f8')

        # タイトル
        tk.Label(
            root, text='案分析レポート生成ツール',
            font=('Helvetica', 18, 'bold'), bg='#f0f4f8', fg='#1a365d',
        ).pack(pady=(24, 8))

        tk.Label(
            root, text='時間枠一覧のExcelファイルを選択してください',
            font=('Helvetica', 11), bg='#f0f4f8', fg='#4a5568',
        ).pack(pady=(0, 16))

        # ファイル選択エリア
        frame = tk.Frame(root, bg='#f0f4f8')
        frame.pack(padx=24, fill='x')

        self.file_var = tk.StringVar(value='ファイルが選択されていません')
        tk.Label(
            frame, textvariable=self.file_var,
            font=('Helvetica', 10), bg='white', fg='#2d3748',
            relief='sunken', anchor='w', padx=8, pady=6,
        ).pack(side='left', fill='x', expand=True)

        tk.Button(
            frame, text='選択', font=('Helvetica', 10, 'bold'),
            bg='#4472C4', fg='white', padx=12, pady=4,
            command=self.select_file, cursor='hand2',
        ).pack(side='right', padx=(8, 0))

        # 実行ボタン
        self.run_btn = tk.Button(
            root, text='レポート生成', font=('Helvetica', 13, 'bold'),
            bg='#2b6cb0', fg='white', padx=32, pady=8,
            command=self.run, cursor='hand2', state='disabled',
        )
        self.run_btn.pack(pady=24)

        # ステータス
        self.status_var = tk.StringVar(value='')
        self.status_label = tk.Label(
            root, textvariable=self.status_var,
            font=('Helvetica', 10), bg='#f0f4f8', fg='#2d3748',
            wraplength=460,
        )
        self.status_label.pack(padx=24, pady=(0, 8))

        # 結果表示
        self.result_var = tk.StringVar(value='')
        tk.Label(
            root, textvariable=self.result_var,
            font=('Helvetica', 10, 'bold'), bg='#f0f4f8', fg='#276749',
            wraplength=460,
        ).pack(padx=24)

        self.input_path = None

    def select_file(self):
        path = filedialog.askopenfilename(
            title='Excelファイルを選択',
            filetypes=[('Excel files', '*.xlsx *.xlsm'), ('All files', '*.*')],
        )
        if path:
            self.input_path = path
            self.file_var.set(os.path.basename(path))
            self.run_btn.configure(state='normal')
            self.status_var.set('')
            self.result_var.set('')

    def run(self):
        if not self.input_path:
            return
        self.run_btn.configure(state='disabled')
        self.status_var.set('処理中...')
        self.result_var.set('')
        threading.Thread(target=self._generate, daemon=True).start()

    def _generate(self):
        try:
            output_path = generate_report(self.input_path)
            self.root.after(0, self._on_success, output_path)
        except Exception as e:
            self.root.after(0, self._on_error, str(e))

    def _on_success(self, output_path):
        self.status_var.set('完了!')
        self.result_var.set(f'保存先: {output_path}')
        self.run_btn.configure(state='normal')
        messagebox.showinfo('完了', f'レポートを生成しました。\n\n{output_path}')

    def _on_error(self, msg):
        self.status_var.set(f'エラー: {msg}')
        self.result_var.set('')
        self.run_btn.configure(state='normal')
        messagebox.showerror('エラー', msg)


def main():
    # コマンドライン引数がある場合はCLIモードで実行
    if len(sys.argv) >= 2:
        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) >= 3 else None
        generate_report(input_file, output_file)
        return

    # GUI起動
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
