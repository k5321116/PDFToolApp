import os
import glob
import win32com.client  # pywin32 ライブラリ
import time

def convert_word_to_pdf(source_folder, dest_folder):
    """
    指定されたフォルダー(source_folder)内のWordファイルをすべてPDFに変換し、
    指定された保存先フォルダー(dest_folder)に保存します。
    """
    
    # 1. Wordアプリケーションを起動
    try:
        word_app = win32com.client.Dispatch("Word.Application")
    except Exception as e:
        print(f"Wordの起動に失敗しました。Wordがインストールされているか確認してください。")
        print(f"エラー: {e}")
        return

    word_app.Visible = False  # Wordを画面に表示せずに実行
    WD_FORMAT_PDF = 17        # PDF保存時のフォーマット番号

    try:
        # 2. 宛先フォルダーが存在しない場合は作成
        os.makedirs(dest_folder, exist_ok=True)
        
        # 3. 変換元フォルダー内のWordファイルリストを取得
        word_files = []
        word_files.extend(glob.glob(os.path.join(source_folder, '*.docx')))
        word_files.extend(glob.glob(os.path.join(source_folder, '*.doc')))

        if not word_files:
            print(f"変換元フォルダー内にWordファイルが見つかりません: {source_folder}")
            return

        print(f"{len(word_files)}個のファイルを変換します...")

        for raw_word_path in word_files:
            # ★修正ポイント: パスをWindows形式の絶対パスに整形する (混在した / を \ に直す)
            word_path_abs = os.path.abspath(raw_word_path)
            
            # PDFのファイル名作成
            base_filename = os.path.basename(word_path_abs) 
            pdf_filename = os.path.splitext(base_filename)[0] + '.pdf'
            
            # PDFのパスも整形する
            raw_pdf_path = os.path.join(dest_folder, pdf_filename)
            pdf_path_abs = os.path.abspath(raw_pdf_path)
            
            print(f"  変換中: {base_filename} -> {pdf_path_abs}")

            # 4. Wordドキュメントを開く
            try:
                doc = word_app.Documents.Open(word_path_abs)
                
                # 5. PDFとして保存
                doc.SaveAs(pdf_path_abs, FileFormat=WD_FORMAT_PDF)
                
                # 6. ドキュメントを閉じる
                doc.Close()
            except Exception as e:
                print(f"  ※ファイル '{base_filename}' の処理中にエラー: {e}")
                # エラーが出ても次のファイルの処理を続けるため continue しない（ループは継続）

        print("すべての処理が完了しました。")

    except Exception as e:
        print(f"変換処理全体でエラーが発生しました: {e}")
    
    finally:
        # 7. Wordアプリケーションを終了
        word_app.Quit()

# --- 実行 ---
if __name__ == "__main__":
    
    # 1. 変換元のWordファイルが入っているフォルダーのパス
    source_folder_path = "C:/Users/k5321_ovwb2/OneDrive/デスクトップ/wordfiles"
    
    # 2. 変換後のPDFを入れるフォルダーのパス
    destination_folder_path = "C:/Users/k5321_ovwb2/OneDrive/デスクトップ/files"
    
    # --- (これ以降は変更不要) ---
    
    if not os.path.isdir(source_folder_path):
        print(f"エラー: 変換元のフォルダーが見つかりません。")
        print(f"パス: {source_folder_path}")
    else:
        # 念のため入力されたパス自体も整形して表示確認
        print(f"変換元: {os.path.abspath(source_folder_path)}")
        print(f"保存先: {os.path.abspath(destination_folder_path)}")
        
        convert_word_to_pdf(source_folder_path, destination_folder_path)