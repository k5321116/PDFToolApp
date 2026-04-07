import os
import io
import sys
import time
import glob
import win32com.client
import pythoncom
from multiprocessing import Pool, freeze_support

if sys.stdout is not None:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def convert_single_file_robust(args):
    word_path_abs, dest_folder = args
    pythoncom.CoInitialize()
    
    word_app = None
    max_retries = 3
    file_start = time.time()
    
    success = False
    for i in range(max_retries):
        try:
            word_app = win32com.client.DispatchEx("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0
            
            pdf_filename = os.path.splitext(os.path.basename(word_path_abs))[0] + '.pdf'
            pdf_path_abs = os.path.abspath(os.path.join(dest_folder, pdf_filename))
            
            doc = word_app.Documents.Open(word_path_abs, ReadOnly=True)
            doc.SaveAs(pdf_path_abs, FileFormat=17)
            doc.Close(0)
            success = True
            break
        except Exception:
            if word_app:
                try: word_app.Quit()
                except: pass
            time.sleep(1)
        finally:
            if word_app:
                try: word_app.Quit()
                except: pass

    pythoncom.CoUninitialize()
    return (success, time.time() - file_start)

def run_conversion(source_folder, dest_folder):
    # globで探すのは source_folder
    all_files = [os.path.abspath(f) for f in glob.glob(os.path.join(source_folder, '*.doc*'))]
    word_files = [f for f in all_files if not os.path.basename(f).startswith('~$')]
    
    if not word_files:
        print("有効なWordファイルが見つかりませんでした。", flush=True)
        return

    total = len(word_files)
    finished = 0
    # 変換関数に dest_folder を渡す
    task_args = [(f, dest_folder) for f in word_files]
    num_processes = 2 
    
    os.system("taskkill /f /im WINWORD.EXE >nul 2>&1")

    with Pool(processes=num_processes) as pool:
        for result in pool.imap_unordered(convert_single_file_robust, task_args):
            finished += 1
            percent = int((finished / total) * 100)
            # JS側が受け取れる形式で出力
            print(f"PROGRESS:{percent}", flush=True)
    print(f"【完了】{total}件の処理に成功しました。", flush=True)

if __name__ == "__main__":
    freeze_support()  # Windowsでのマルチプロセッシングに必要
    # 引数が2つあるか確認
    if len(sys.argv) >= 3:
        run_conversion(sys.argv[1], sys.argv[2])
    else:
        print("エラー: フォルダパスが正しく指定されていません。", flush=True)