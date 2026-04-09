# PDF Utility Suite

Word ファイルの PDF 変換と、複数 PDF の結合をワンストップで行える Windows 向けデスクトップアプリです。  
Electron + Python（PyInstaller でビルドした exe）で動作します。

---

## 主な機能

| 機能 | 説明 |
|---|---|
| Word → PDF 変換 | 指定フォルダ内の `.doc` / `.docx` を一括変換。並列処理（2プロセス）で高速化。 |
| PDF 結合 | 複数の PDF ファイルをドラッグ＆ドロップで並べ替えて結合。 |
| ページ番号挿入 | 結合時にページ番号を自動付与（特定ページ・特定ファイルは除外可）。 |
| ドラッグ＆ドロップ | フォルダ／ファイルのドロップに対応。 |
| 削除済みファイルの復元 | 結合リストから外したファイルをゴミ箱から復元できる。 |

---

## 動作環境

- **OS**: Windows 10 / 11
- **Microsoft Word**: インストール済みであること（Word → PDF 変換機能を使う場合）
- **Node.js**: 開発時のみ必要（v18 以上推奨）
- **Python**: 開発時のみ必要（v3.9 以上推奨）

---

## フォルダ構成

```
project/
├── main.js              # Electron メインプロセス
├── index.html           # UI（レンダラープロセス）
├── package.json
└── python/
    ├── Word_to_PDF.py   # Word → PDF 変換スクリプト（開発時）
    ├── Word_to_PDF.exe  # 〃 ビルド済み（配布時）
    ├── pdf_merger.py    # PDF 結合スクリプト（開発時）
    └── pdf_merger.exe   # 〃 ビルド済み（配布時）
```

---

## セットアップ（開発環境）

### 1. Node.js パッケージのインストール

```bash
npm install
```

### 2. Python ライブラリのインストール

```bash
pip install pywin32 pypdf reportlab
```

### 3. 起動

```bash
npm start
```

---

## Python スクリプトを exe にビルドする

配布用の exe は PyInstaller で作成します。

### Word_to_PDF.exe

```bash
pyinstaller --onefile --noconsole ^
  --hidden-import=win32com ^
  --hidden-import=win32com.client ^
  --hidden-import=pythoncom ^
  --hidden-import=pywintypes ^
  Word_to_PDF.py
```

### pdf_merger.exe

```bash
pyinstaller --onefile --noconsole pdf_merger.py
```

ビルド後、`dist/` に生成された exe を `python/` フォルダへ移動してください。

> **注意**: `Word_to_PDF.py` は `multiprocessing` を使用しているため、  
> `if __name__ == "__main__":` ブロック内の `freeze_support()` が必須です。  
> これがないと exe 実行時にクラッシュします。

---

## 使い方

### Word → PDF 変換

1. ホーム画面で「Word → PDF 変換」を選択
2. 変換元フォルダを選択（またはドロップ）
3. 保存先フォルダを選択
4. 「変換開始」をクリック
5. 完了後、「PDF 結合画面へ」ボタンでそのまま結合作業に移れます

### PDF 結合

1. ホーム画面で「PDF 結合」を選択
2. ファイル・フォルダを追加（またはドロップ）→「次へ」
3. ドラッグ＆ドロップで順序を調整
4. オプションを設定：
   - ページ番号を挿入するか
   - ページ番号を除外するページ番号（例: `1, 8`）
   - ページ番号を付けないファイル（各ファイル行のチェックボックスで制御）
5. 保存先ファイル名を指定 → 「結合開始」

---

## トラブルシューティング

### exe にすると動かない

`main.js` では `app.isPackaged` フラグで開発時と配布時を切り替えています。  
配布時は `python-shell` ではなく `child_process.spawn` で exe を直接起動します。  
以下を確認してください：

- `python/Word_to_PDF.exe` と `python/pdf_merger.exe` が正しい場所に存在するか
- Electron Builder の `extraResources` に `python/` フォルダが含まれているか

### Word 変換でエラーになる

- Microsoft Word がインストールされているか確認してください
- 変換前に Word が起動したままになっている場合、自動で `WINWORD.EXE` を強制終了してから処理を開始します
- それでも失敗する場合は、手動で Word を終了してから再試行してください

### 文字化けする

Python スクリプト側で `sys.stdout` を UTF-8 に設定済みです。  
環境によって問題が出る場合は、システムの「地域の設定」→「Unicode UTF-8 を使用」を有効にしてください。

---

## ライセンス

MIT License
