# book_kdl - Gemini Automation

Google Drive上のDOCXファイルをGeminiで処理し、結果をスプレッドシートに集約する自動化ツール。

## セットアップ

### 1. 依存ライブラリのインストール

```bash
pip install -r requirements.txt
```

または個別に:

```bash
pip install pyautogui pywinauto Pillow opencv-python
```

### 2. 設定ファイルの編集

`config.py` を編集して、環境に合わせてパスを設定:

```python
# Google DriveのDOCXファイルがあるフォルダ
GOOGLE_DRIVE_BOOK_KDL_PATH = "G:/My Drive/book_kdl"

# 集計先のスプレッドシート名
MASTER_SPREADSHEET_NAME = "KDL_20260115"
```

### 3. ボタン画像のキャプチャ（オプション）

画像認識ベースの検出を使用する場合:

```bash
python capture_buttons.py
```

以下のボタン画像をキャプチャ:
- `gem_button.png` - Gem選択ボタン
- `upload_button.png` - ファイルアップロード（+）ボタン
- `send_button.png` - 送信ボタン
- `spreadsheet_button.png` - スプレッドシートを開くボタン

## 実行方法

### 事前準備

1. ブラウザでGeminiを開いておく
2. マスタースプレッドシート（`KDL_20260115`）を開いておく
3. Google DriveにDOCXファイルを配置

### 実行

```bash
python gemini_automation.py
```

## ファイル構成

```
book_kdl/
├── config.py              # 設定ファイル
├── gemini_automation.py   # メイン実行スクリプト
├── gemini_ops.py          # Gemini操作モジュール
├── ui_helpers.py          # UI操作ユーティリティ
├── capture_buttons.py     # ボタン画像キャプチャツール
├── requirements.txt       # 依存ライブラリ
├── images/                # ボタン画像（自動生成）
└── design/
    └── automation_plan.md # 設計書
```

## 処理フロー

1. Google Drive内のDOCXファイル一覧を取得
2. 各ファイルに対して:
   - Geminiウィンドウをアクティブ化
   - Gemを選択してプロンプト入力
   - ファイルをアップロード
   - 実行して応答を待機
   - 生成されたスプレッドシートを開く
   - データをコピー
   - マスターシートに貼り付け
3. 完了サマリーを表示

## 緊急停止

マウスを画面の隅に素早く移動させると、pyautoguiのフェイルセーフが作動して停止します。

## 注意事項

- 処理中はマウス・キーボードを操作しないでください
- Geminiの応答時間によっては待機時間の調整が必要な場合があります
- スプレッドシートはGoogle Sheets形式を想定しています
