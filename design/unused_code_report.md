# 削除体操（削除候補コード一覧）

現在の `gemini_automation.py` は、ハードコードされた座標と `pyautogui` を直接使用する手続き型処理に変更されています。その結果、以下のクラスメソッドは現在使用されておらず、削除候補となります。

## 1. gemini_ops.py

### Class: GeminiAutomation
以下のメソッドは `process_single_file` 内でコメントアウトされ、直接的な座標クリック操作に置き換えられています。

- `input_prompt`
- `click_upload1`
- `click_upload2`
- `upload_file_dialog`
- `execute_prompt`
- `wait_for_completion`
- `open_spreadsheet`
- `confirm_spreadsheet_dialog`

*注: `find_gemini_window`, `activate_gemini`, `select_gem` はまだ使用されています。*

### Class: SpreadsheetAutomation
以下のメソッドは現在使用されていません（スプレッドシート操作も座標クリックとショートカットキーで行われています）。

- `find_spreadsheet_window` (直接呼び出されていないが `activate_master_sheet` から呼ばれる)
- `copy_data`
- `activate_master_sheet`
- `paste_data`
- `close_current`

## 2. gemini_automation.py

### Global Functions
- `get_docx_files`: 内部で `glob` を使用しているが、`main` 関数内ではハードコードされたファイルリストが使用されているため、実質的に未使用（またはデバッグ用）。

---
**推奨アクション**:
現在の「座標指定型」の自動化を維持する場合、これらのメソッドはコードベースを肥大化させるだけのため、削除または `legacy_ops.py` 等へ待避することを推奨します。
