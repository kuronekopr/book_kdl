# 詳細設計書: Gemini 自動化モジュール (Coordinate-Based Implementation)

本ドキュメントは、現在運用されている `gemini_automation.py` および `gemini_ops.py` の実装状態を反映した詳細設計書です。
初期の画像認識ベースの設計から、座標指定ベースの手続き型実装へと移行しています。

## 1. 概要
本システムは、ハードコードされた画面座標と `pyautogui` を使用して、Google Gemini Web インターフェースを操作し、ドキュメント処理を行います。
処理は `Google Drive/My Drive/book_kdl` フォルダ内の DOCX ファイルリスト（ハードコード定義）に対して順次実行されます。

## 2. モジュール構成
- **メインスクリプト**: `gemini_automation.py`
  - 実行エントリーポイント (`main`)
  - 座標定義
  - 手続き型処理フロー (`process_single_file`)
- **操作ライブラリ**: `gemini_ops.py`
  - 初期のクラス定義 (`GeminiAutomation`, `SpreadsheetAutomation`) が残存しているが、大半のメソッドは未使用。
  - `activate_gemini` や `select_gem` など一部のウィンドウ制御系メソッドのみ使用。

## 3. 定数定義 (座標設定)
`gemini_automation.py` 内で定義されている主要な操作座標です。環境依存性が高いため、解像度やウィンドウ配置が変わる場合は再調整が必要です。

| 変数名 | 座標 (x, y) | 説明 |
| :--- | :--- | :--- |
| `upload_butto1_x/y` | (675, 1942) | アップロードダイアログ呼び出し等のボタン |
| `upload_butto2_x/y` | (730, 1731) | 予備/追加のアップロード操作点 |
| `gemi_send_x/y` | (1717, 1940) | Gemini プロンプト送信ボタン |
| `ans_focus_x/y` | (1162, 439) | 回答エリアへのフォーカス |
| `sp_sel1` ～ `sp_sel6` | (各所) | スプレッドシート生成オプションやチェックボックスの選択操作 |
| `kdl_focus_x/y` | (896, 34) | マスターファイル (KDL_20260115) ウィンドウのフォーカス位置 |
| `sp_close_x/y` | (738, 21) | 生成された一時スプレッドシートタブの閉じるボタン |
| `gem_focus_x/y` | (87, 20) | Gemini タブへの復帰フォーカス |

## 4. 処理詳細フロー (`gemini_automation.py`)

### メインルーチン (`main`)
1. **初期化**: `GeminiAutomation`, `SpreadsheetAutomation` のインスタンス化。
2. **対象リスト定義**: DOCXファイル名のハードコードリストを使用（`get_docx_files` は使用停止）。
3. **ループ処理**: 各ファイルについて `process_single_file` を呼び出す。

### ファイル処理ルーチン (`process_single_file`)
各ファイルに対して以下のステップを順次実行します。待機は `time.sleep()` で固定時間制御されています。

1. **準備**
   - ファイル名をクリップボードにコピー (`clipboard.copy`)。
   - `gemini.activate_gemini()`: Geminiウィンドウをアクティブ化。
   - `gemini.select_gem()`: Gemを選択。

2. **アップロードと入力** (座標操作)
   - `upload_butto1` クリック。
   - `upload_butto2` クリック (ファイル選択メニュー操作と推測)。
   - `Ctrl+V`: クリップボードのファイル名を貼り付け。
   - `Enter`: ファイル確定。
   - **待機 (10秒)**: アップロード完了待ち。
   - `gemi_send`: 送信ボタンクリック。

3. **実行と待機**
   - **待機 (120秒)**: Geminiの生成処理およびスプレッドシート作成待ち。
   - `ans_focus`: 回答エリアをクリック。
   - `End` キー: 画面下部へスクロール。

4. **結果抽出 (スプレッドシート操作)**
   - `sp_sel1` ～ `sp_sel6` を順次クリック: 生成オプションの選択やダウンロード設定等。
   - `Ctrl+A`: 全選択。
   - `Ctrl+C`: コピー。

5. **マスタースプレッドシートへの追記**
   - `kdl_focus`: マスターファイルウィンドウをクリックしてアクティブ化。
   - `Ctrl+Home`: データの先頭へ。
   - `Ctrl+Down` x 2: データの末尾周辺へ移動 (既存データのスキップ)。
   - `Down`: 新規行へ移動。
   - `Ctrl+Shift+V`: 値のみ貼り付け。

6. **終了処理**
   - `sp_close`: 一時タブを閉じる。
   - `gem_focus`: Geminiタブへ戻る。

## 5. 既存コードとの乖離 (Unused Code)
以下のクラスベースメソッドは、現在の `process_single_file` 手続き内では呼び出されていません (詳細は `unused_code_report.md` 参照)。
- `GeminiAutomation` の `click_upload`, `input_prompt`, `execute_prompt` 等
- `SpreadsheetAutomation` の全メソッド (`copy_data`, `paste_data` 等)

## 6. 注意事項
- **座標依存**: 現在の実装は特定の画面解像度とウィンドウ位置に完全に依存しています。ウィンドウサイズ変更やディスプレイ変更で動作しなくなります。
- **エラー処理**: 画像認識による状態確認（`wait_for_image`）が行われていないため、タイムアウトや予期せぬポップアップに弱くなっています。
