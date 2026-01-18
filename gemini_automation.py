"""
Gemini Automation Script
Automates document processing through Gemini Web interface
and aggregates results to a master spreadsheet.
"""

import os
import sys
import time

import pyautogui
import clipboard

from config import (
    GOOGLE_DRIVE_BOOK_KDL_PATH,
    MASTER_SPREADSHEET_NAME,
    WAIT_MEDIUM,
    WAIT_SHORT,
    WAIT_SPREADSHEET
)
from gemini_ops import GeminiAutomation, log


# Fail-safe: move mouse to corner to abort
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5

upload_butto1_x = 675
upload_butto1_y = 1942
upload_butto2_x = 730
upload_butto2_y = 1731
gemi_send_x = 1717
gemi_send_y = 1940
ans_focus_x = 1162
ans_focus_y = 439
sp_sel1_x = 950
sp_sel1_y = 1470
sp_sel2_x = 950
sp_sel2_y = 1420
sp_sel3_x = 950
sp_sel3_y = 1505
sp_sel4_x = 950
sp_sel4_y = 1520
sp_sel5_x = 950
sp_sel5_y = 1615
sp_sel6_x = 579
sp_sel6_y = 2010
kdl_focus_x = 896
kdl_focus_y = 34
sp_close_x = 738
sp_close_y = 21
gem_focus_x = 87
gem_focus_y = 20





def process_single_file(
    file_path: str,
    gemini: GeminiAutomation
) -> bool:
    """Process a single DOCX file through Gemini."""
    log(f"\n{'='*50}")
    log(f"Processing: {os.path.basename(file_path)}")
    log(f"{'='*50}")

    clipboard.copy(os.path.basename(file_path))
    try:
        # Step 1: Ensure Gemini is active
        if not gemini.activate_gemini():
            return False

        # Step 2: Select Gem
        gemini.select_gem()

        # Step 3: Input prompt
        #gemini.input_prompt()

        # Step 3: Upload file
        #gemini.click_upload1()
        #time.sleep(WAIT_MEDIUM)
        #gemini.upload_file_dialog(file_path)
        time.sleep(WAIT_SHORT)
        pyautogui.click(x=upload_butto1_x, y=upload_butto1_y)
        time.sleep(WAIT_SHORT)
        pyautogui.click(x=upload_butto2_x, y=upload_butto2_y)
        time.sleep(WAIT_MEDIUM)
        pyautogui.hotkey('ctrl','v')
        time.sleep(WAIT_SHORT)
        pyautogui.press('enter')
        time.sleep(10)
        pyautogui.click(x=gemi_send_x, y=gemi_send_y)
        time.sleep(120)
        pyautogui.click(x=ans_focus_x, y=ans_focus_y)
        time.sleep(WAIT_SHORT)
        pyautogui.press('end')
        time.sleep(WAIT_SHORT)
        
        pyautogui.click(x=sp_sel1_x, y=sp_sel1_y)
        time.sleep(WAIT_SHORT)
        pyautogui.click(x=sp_sel2_x, y=sp_sel2_y)
        time.sleep(WAIT_SHORT)
        pyautogui.click(x=sp_sel3_x, y=sp_sel3_y)
        time.sleep(WAIT_SHORT)
        pyautogui.click(x=sp_sel4_x, y=sp_sel4_y)
        time.sleep(WAIT_SHORT)
        pyautogui.click(x=sp_sel5_x, y=sp_sel5_y)
        time.sleep(WAIT_MEDIUM)
        pyautogui.click(x=sp_sel6_x, y=sp_sel6_y)
        time.sleep(WAIT_MEDIUM)
        pyautogui.hotkey('ctrl','a')
        #pyautogui.hotkey('ctrl','shift','down',wait_after=WAIT_SHORT)
        #time.sleep(WAIT_SHORT)
        #pyautogui.hotkey('right',wait_after=WAIT_SHORT)
        #time.sleep(WAIT_SHORT)
        #pyautogui.hotkey('ctrl','shift','right')
        time.sleep(WAIT_SHORT)
        pyautogui.hotkey('ctrl','c')
        time.sleep(WAIT_SHORT)
        pyautogui.click(x=kdl_focus_x, y=kdl_focus_y)
        time.sleep(WAIT_SHORT)
        pyautogui.hotkey('ctrl','home')
        time.sleep(WAIT_SHORT)
        pyautogui.hotkey('ctrl','down')
        time.sleep(WAIT_SHORT)
        pyautogui.hotkey('down')
        time.sleep(WAIT_SHORT)
        pyautogui.hotkey('ctrl','shift','v')
        time.sleep(WAIT_SHORT)
        pyautogui.click(x=sp_close_x, y=sp_close_y)
        time.sleep(WAIT_SHORT)
        pyautogui.click(x=gem_focus_x, y=gem_focus_y)

        #gemini.upload_file_dialog(file_path)


        
        # Step 4: Upload file


        # Step 5: Execute
        #gemini.execute_prompt()

        # Step 6: Wait for response
        #gemini.wait_for_completion(timeout=90.0)

        # Step 7: Open spreadsheet
        #gemini.open_spreadsheet()
        #gemini.confirm_spreadsheet_dialog()

        # Step 8: Copy data from generated spreadsheet
        #time.sleep(WAIT_SPREADSHEET)
        #spreadsheet.copy_data()

        # Step 9: Paste to master sheet
        #if not spreadsheet.activate_master_sheet(MASTER_SPREADSHEET_NAME):
        #    log("WARNING: Could not activate master spreadsheet")
        #    return False

        #spreadsheet.paste_data()

        # Step 10: Close temp spreadsheet and return to Gemini
        #spreadsheet.close_current()

        #log(f"Successfully processed: {os.path.basename(file_path)}")
        return True

    except Exception as e:
        log(f"ERROR processing file: {e}")
        return False


def main():
    """Main entry point for the automation."""
    log("=" * 60)
    log("GEMINI AUTOMATION - Document Processing")
    log("=" * 60)
    log(f"Source folder: {GOOGLE_DRIVE_BOOK_KDL_PATH}")
    log(f"Master sheet: {MASTER_SPREADSHEET_NAME}")

    # Verify source path exists
    if not os.path.exists(GOOGLE_DRIVE_BOOK_KDL_PATH):
        log(f"ERROR: Source path does not exist: {GOOGLE_DRIVE_BOOK_KDL_PATH}")
        log("Please update GOOGLE_DRIVE_BOOK_KDL_PATH in config.py")
        sys.exit(1)

    # Get list of files to process
    #files = get_docx_files()
    files = ['THE YOROSHIKU GUIDE _ THE よろしく GUIDE_ Japanese Business Edition.docx','システム導入のためのデータ移行ガイドブック―コンサルタントが現場で体得したデータ移行のコツ (NextPublishing).docx','戦わずして売る技術　クリック１つで市場を生み出す最強のWEBマーケティング術 (幻冬舎単行本)_B0FMQ1Q69F.kfx.docx','マジビジプロ 新人コンサルタントが最初に学ぶ 厳選フレームワーク20 MAJIBIJI pro［図解］問題解決に強くなる！.docx','実践Claude Code入門―現場で活用するためのAIコーディングの思考法 エンジニア選書_B0G4NVX1NF.kfx.docx','Markdownライティング入門　プレーンテキストで気楽に書こう！ (技術の泉シリーズ（NextPublishing）).docx','トリーズ(TRIZ)の発明原理40 あらゆる問題解決に使える[科学的]思考支援ツール トリーズ（TRIZ）の発明原理４０.docx','身体を壊す健康法 年間500本以上読破の論文オタクの東大医学博士＆現役医師が、世界中から有益な情報を見つけて解き明かす。.docx','いちばんやさしい機械学習プロジェクトの教本　人気講師が教える仕事にAIを導入する方法 「いちばんやさしい教本」シリーズ.docx','視点転換転職術_ 転職活動を勝ち進む5つのノウハウ 電子書籍で学ぶ転職 (KBLT出版)_B0FLQ7CDGP.kfx.docx','FIRE生活への近道！正しい副業のはじめ方_ 凡人サラリーマンが貯金ゼロから3年で独立した方法 会社に頼らない生き方.docx','【図解】はじめての上流工程(要件定義・システム設計・プロジェクトマネジメント)入門_ よくわかる！システム開発入門.docx','NotebookLM最強勉強法_ ライバルはまだ知らないAIで差をつける最新勉強法_B0FSJTGF4B.kfx.docx','PLG プロダクト・レッド・グロース「セールスがプロダクトを売る時代」から「プロダクトでプロダクトを売る時代」へ.docx','マッキンゼーで叩き込まれた 超速フレームワーク―――仕事のスピードと質を上げる最強ツール (三笠書房　電子書籍).docx','外資系コンサルのPowerPoint資料作成140のテクニック_ スライドをわかりやすく、速くつくる理論と技術.docx','タロットと遊んで未来をつくる方法_ リアル宝探しゲーム編 タロットで未来創造シリーズ (占わない師の道具箱).docx','Claude-3.7 完全活用ガイド_ AI初心者でも今日から始められる業務効率化と収益化の実践バイブル.docx','High Conflict_ Why We Get Trapped and How We Get Out.docx','良いFAQの書き方──ユーザーの「わからない」を解決するための文章術 WEB+DB PRESS plus.docx','Microsoft Wordを開発した伝説のプログラマーが発見した「やりたいことの見つけ方」がすごい！.docx','Risk Up Front_ Managing Projects in a Complex World.docx','チャートで考えればうまくいく 一生役立つ「構造化思考」養成講座 (セブンチャートテンプレート特典付き).docx','FACTFULNESS（ファクトフルネス）10の思い込みを乗り越え、データを基に世界を正しく見る習慣.docx','アジャイルデータモデリング　組織にデータ分析を広めるためのテーブル設計ガイド (ＫＳ情報科学専門書).docx','結果を出すクレーム。店、企業、役所、病院、学校、職場、もう泣き寝入りしない！30分で読めるシリーズ.docx','生成AI×ビジネススキル_ 【仕事術】【効率化】【タイパ】【マーケティング】【セールス】【差別化】.docx','The Intelligent Sales AIを活用した最速・最良でクリエイティブな営業プロセス.docx','Vibe Coding for Beginners with Python and ChatGPT.docx','今日がもっと楽しくなる行動最適化大全　ベストタイムにベストルーティンで常に「最高の1日」を作り出す.docx','現場で活用するためのＡＩエージェント実践入門 (ＫＳ情報科学専門書)_B0FM3QJ2DP.kfx.docx','LangChainとLangGraphによるRAG・AIエージェント［実践］入門 エンジニア選書.docx','シン・ニホン AI×データ時代における日本の再生と人材育成 (NewsPicksパブリッシング).docx','今すぐ始められる「ショート動画編集」の教科書　この一冊でOK！未経験から最速で月10万円稼ぐ方法.docx','実践Androidアプリシステムテスト　実機でテスト！チームで挑む品質向上入門 技術の泉シリーズ.docx','これからの時代は１人で自動化で稼ぎなさい！_ 寝ている間も稼ぐ！1人社長の完全自動化戦略とは？.docx','AI時代の質問力 プロンプトリテラシー 「問い」と「指示」が生成AIの可能性を最大限に引き出す.docx','いつまで英語から逃げてるの？ 英語の多動力New Version 君の未来を変える英語のはなし.docx','エンジニアのためのWord再入門講座 新版 美しくメンテナンス性の高い開発ドキュメントの作り方.docx','タロット心理学の入門書_ タロット占いをセラピーレベルに引き上げるスピリチュアルと学問の交差点.docx','なぜあの人の解決策はいつもうまくいくのか？―小さな力で大きく動かす！システム思考の上手な使い方.docx','マッキンゼーのエリートはノートに何を書いているのか　トップコンサルタントの考える技術・書く技術.docx','もっと早く、もっと楽しく、仕事の成果をあげる法 知恵がどんどん湧く「戦略的思考力」を身につけよ.docx','仕事選びのアートとサイエンス～不確実な時代の天職探し　改訂『天職は寝て待て』～ (光文社新書).docx','成功率を高める新規事業のつくり方 顧客・機能・技術の3軸で新しいビジネスアイデアを創出する方法.docx','全米ナンバーワンビジネススクールで教える起業家の思考と実践術―あなたも世界を変える起業家になる.docx','１分で話せ２【超実践編】　世界のトップが絶賛した即座に考えが“まとまる”“伝わる”すごい技術.docx','Python FlaskによるWebアプリ開発入門 物体検知アプリ&機械学習APIの作り方.docx','ソフトウェア開発現場の「失敗」集めてみた。 42の失敗事例で学ぶチーム開発のうまい進めかた.docx','データ可視化の基本が全部わかる本 収集・変換からビジュアライゼーション・データ分析支援まで.docx','トーキング・トゥ・ストレンジャーズ～「よく知らない人」について私たちが知っておくべきこと～.docx','最高の打ち手が見つかるマーケティングの実践ガイド 3つのマップで戦略に沿った施策を実行する.docx','生成AIが起こすビジネス変革_ 未来のビジネスを予測する方法 Kindle生成AI解説本.docx','外資系コンサルのビジネス文書作成術―ロジカルシンキングと文章術によるＷｏｒｄ文書の作り方.docx','現場で使える！pandasデータ前処理入門 機械学習・データサイエンスで役立つ前処理手法.docx']
    if not files:
        log("No .docx files found to process")
        sys.exit(0)

    # Initialize automation handlers
    gemini = GeminiAutomation()
    gemini = GeminiAutomation()

    # Process each file
    successful = 0
    failed = 0
    failed_files = []

    for i, file_path in enumerate(files):
        log(f"\nProgress: {i+1}/{len(files)}")

        if process_single_file(file_path, gemini):
            successful += 1
        else:
            failed += 1
            failed_files.append(os.path.basename(file_path))

        # Small delay between files
        time.sleep(WAIT_MEDIUM)

    # Summary
    log("\n" + "=" * 60)
    log("AUTOMATION COMPLETE")
    log("=" * 60)
    log(f"Successful: {successful}")
    log(f"Failed: {failed}")
    log(f"Total: {len(files)}")

    if failed_files:
        log("\nFailed files:")
        for f in failed_files:
            log(f"  - {f}")

    log("=" * 60)


if __name__ == "__main__":
    main()



