import sys, io, os, re
from pypdf import PdfWriter, PdfReader
from reportlab.pdfgen import canvas

# 日本語パス・文字化け対策
if sys.stdout is not None:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def create_page_number_pdf(num, total, width, height):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(width, height))
    can.setFont("Helvetica", 10)
    text = f"{num}"
    tw = can.stringWidth(text, "Helvetica", 10)
    can.drawString((width - tw) / 2, 20, text)
    can.save()
    packet.seek(0)
    return packet

# pdf_merger.py の修正版

def merge_pdfs(paths_str, save_path, add_page_num, exclude_pages_str, exclude_files_str):
    try:
        exclude_pages = [int(x.strip()) for x in exclude_pages_str.split(',') if x.strip().isdigit()]
        # 除外ファイルパスのリスト
        exclude_files = [p for p in exclude_files_str.split('|') if p]

        file_list = [p for p in paths_str.split('|') if p]
        writer = PdfWriter()
        
        all_pages_info = []
        for p in file_list:
            reader = PdfReader(p)
            for page in reader.pages:
                all_pages_info.append({"page": page, "source": p})
        
        total = len(all_pages_info)
        for i, info in enumerate(all_pages_info):
            page_num = i + 1
            page = info["page"]
            source = info["source"]
            
            # 除外ファイル判定（パスが一致するか）
            is_excluded_file = False
            for ex in exclude_files:
                if os.path.exists(source) and os.path.exists(ex):
                    if os.path.samefile(source, ex):
                        is_excluded_file = True
                        break

            # 合成判定
            if add_page_num and (page_num not in exclude_pages) and not is_excluded_file:
                w, h = float(page.mediabox.width), float(page.mediabox.height)
                num_pdf = PdfReader(create_page_number_pdf(page_num, total, w, h))
                page.merge_page(num_pdf.pages[0])
            
            writer.add_page(page)
            print(f"PROGRESS:{int((page_num/total)*100)}", flush=True)

        with open(save_path, "wb") as f:
            writer.write(f)
        print(f"✅ 結合完了！", flush=True)
    except Exception as e:
        print(f"Pythonエラー: {str(e)}", flush=True)

if __name__ == "__main__":
    # 引数が5つ（スクリプト名除いて）あるか確認
    if len(sys.argv) >= 6:
        merge_pdfs(sys.argv[1], sys.argv[2], sys.argv[3]=="True", sys.argv[4], sys.argv[5])