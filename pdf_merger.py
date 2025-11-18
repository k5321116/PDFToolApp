import os
import re  # これを追加！

from pypdf import PdfWriter  # pyright: ignore[reportMissingImports]

# フォルダ設定
input_folder = r"C:\Users\k5321_ovwb2\OneDrive\デスクトップ\pdffiles"
output_file = r"C:\Users\k5321_ovwb2\OneDrive\デスクトップ\pdffiles\merged.pdf"

# 出力先フォルダを作成（なければ）
os.makedirs(os.path.dirname(output_file), exist_ok=True)

# PDF結合処理
writer = PdfWriter()
pdf_count = 0

# ソート用の関数を定義（数字を含む場合のため）
def natural_sort_key(file_name):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', file_name)]

# ソート後にPDFを結合
for file_name in sorted(os.listdir(input_folder), key=natural_sort_key):
    if file_name.endswith(".pdf"):
        full_path = os.path.join(input_folder, file_name)
        with open(full_path, "rb") as f:
            writer.append(f)
        pdf_count += 1

# 書き出し
with open(output_file, "wb") as f_out:
    writer.write(f_out)

if pdf_count > 0:
    print(f"✅ PDF {pdf_count}個 結合完了！")
else:
    print("⚠️ 結合対象のPDFがありませんでした（空のPDFを作成）")
