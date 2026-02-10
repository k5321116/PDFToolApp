import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from pypdf import PdfWriter

def natural_sort_key(file_name):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', file_name)]

def merge_pdfs():
    # GUIウィンドウを隠す設定
    root = tk.Tk()
    root.withdraw()
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    
    # 1. 結合したいPDFファイルを複数選択
    files = filedialog.askopenfilenames(
        title="結合するPDFファイルを選択してください",
        filetypes=[("PDF files", "*.pdf")],
        initialdir= desktop_path # 初期ディレクトリ
    )

    if not files:
        print("キャンセルされました")
        return

    # 2. 保存先を指定
    save_path = filedialog.asksaveasfilename(
        title="保存先を選択してください",
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        initialfile="merged.pdf"
    )

    if not save_path:
        return

    # 3. 結合処理
    writer = PdfWriter()
    
    # 選択されたファイルを自然順でソート（ファイル名部分で判定）
    sorted_files = sorted(files, key=lambda x: natural_sort_key(os.path.basename(x)))

    try:
        for path in sorted_files:
            writer.append(path)
        
        with open(save_path, "wb") as f_out:
            writer.write(f_out)
        
        messagebox.showinfo("完了", f"✅ PDF {len(sorted_files)}個を結合しました！\n保存先: {save_path}")
    
    except Exception as e:
        messagebox.showerror("エラー", f"結合中にエラーが発生しました:\n{e}")

if __name__ == "__main__":
    merge_pdfs()