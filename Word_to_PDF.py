import os
import glob
import win32com.client
import tkinter as tk
from tkinter import filedialog, messagebox

def convert_word_to_pdf():
    # GUIウィンドウの設定
    root = tk.Tk()
    root.withdraw()

    # デスクトップのパスを取得
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

    # 1. Wordファイルが入っているフォルダを選択
    source_folder = filedialog.askdirectory(
        title="Wordファイルが入っているフォルダを選択してください",
        initialdir=desktop_path
    )
    if not source_folder: return

    # 2. 保存先フォルダを選択
    dest_folder = filedialog.askdirectory(
        title="PDFの保存先フォルダを選択してください",
        initialdir=source_folder
    )
    if not dest_folder: return

    # Wordアプリケーションの起動
    try:
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
    except Exception as e:
        messagebox.showerror("エラー", f"Wordを起動できませんでした。\n{e}")
        return

    WD_FORMAT_PDF = 17
    word_files = glob.glob(os.path.join(source_folder, '*.doc*')) # doc と docx 両方

    if not word_files:
        messagebox.showwarning("警告", "Wordファイルが見つかりませんでした。")
        word_app.Quit()
        return

    count = 0
    try:
        for word_path in word_files:
            word_path_abs = os.path.abspath(word_path)
            pdf_filename = os.path.splitext(os.path.basename(word_path_abs))[0] + '.pdf'
            pdf_path_abs = os.path.abspath(os.path.join(dest_folder, pdf_filename))

            doc = word_app.Documents.Open(word_path_abs)
            doc.SaveAs(pdf_path_abs, FileFormat=WD_FORMAT_PDF)
            doc.Close()
            count += 1
        
        messagebox.showinfo("完了", f"{count}個のファイルを変換しました！")
    except Exception as e:
        messagebox.showerror("エラー", f"変換中にエラーが発生しました:\n{e}")
    finally:
        word_app.Quit()

if __name__ == "__main__":
    convert_word_to_pdf()