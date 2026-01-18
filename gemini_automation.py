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
    gemini: GeminiAutomation,
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
    files = ['The Two-Second Advantage_ How we succeed by anticipating the future - just enough.docx','The Cold Email Manifesto_ How to fill your sales pipeline, convert like crazy and level up your business in 90 days or less.docx','The Inevitable.docx','【Xと夢の方程式】0からでもSNSで人生は変わる 副業初心者のサラリーマン、主婦がインフル~跡のストーリー　再現性バツグン！無名からX（Twitter）フォロワー1000人までのロードマップ.docx','できる大人のモノの言い方大全.docx','スイッチ！.docx','ゼロからの「TikTokマーケティング」入門 〜フォロワー15万人のインフルエンサーが教える「100万回再生を量産する」極意〜.docx','【Notionの教科書】誰でもできる！スマホとPCで始める新しい自己管理法　心理学とAIを活用し~ド：アイデアの見える化・自動化・目標設定・モチベーション管理で学ぶ成功へのステップ.docx','これからの時代は１人で自動化で稼ぎなさい！_ 寝ている間も稼ぐ！1人社長の完全自動化戦略とは？.docx','【IT × AIで人生設計】未来を作る「欲」の覚醒_ 40代から始める副業・スキル活用で新しい働き方を手に入れる方法.docx','コンサルティングとは何か (PHPビジネス新書).docx','「優良顧客」と「悪質顧客」を100%見抜く方法 —— 「2つ」のパターンに分けるだけ! ストレスフリーの顧客対応マニュアル.docx','『ChatGPT-4o』特級快速AI超実践術_ プロンプト、もはや不要！？図解99枚で5分でマネして即実~基本から実践まで超わかりやすく解説した本【DALL-E3】【画像生成AI】【人工知能】【副業】.docx','フリーズする脳　思考が止まる、言葉に詰まる (生活人新書).docx','【図解の超基本】デザイン初心者でもそれっぽく仕上がる図解の作り方_ 資料作成やプレゼ~月10万円稼ぐ図解デザイナーが徹底解説！見やすいデザインと一生使える図解の黄金ルール.docx','ファスト＆スロー　（下）.docx','【新版】マインドフルネスの教科書 (スピリチュアルの教科書).docx','ファスト＆スロー　（上）.docx','【推しの子】 ～一番星のスピカ～ (ジャンプジェイブックスDIGITAL).docx','【推しの子】 ～二人のエチュード～ (ジャンプジェイブックスDIGITAL).docx','【図解】はじめての上流工程(要件定義・システム設計・プロジェクトマネジメント)入門_ よくわかる！システム開発入門.docx','トレードオフ.docx','10年後に自社が絶滅しない、人財・経営戦略としての「DX」_ 日揮ホールディングスのCDOが明かす「2つの課題」と「3種の神器」.docx','１分で話せ２【超実践編】　世界のトップが絶賛した即座に考えが“まとまる”“伝わる”すごい技術.docx','ルーズヴェルト・ゲーム (講談社文庫).docx','3つのステップで成功させるデータビジネス 「データで稼げる」新規事業をつくる.docx','0→100(ゼロヒャク)生み出す力.docx','リセット　Google流最高の自分を引き出す5つの方法.docx','1秒で心をつかめ。　一瞬で人を動かし、100％好かれる声・表情・話し方.docx','最短１か月で集客！TikTokの専門家が書いたTikTokマーケティングの全て.docx','100%集中法.docx','結局は質問力 ー質問の質で差がつくプロンプトエンジニアリング.docx','結果を出すクレーム。店、企業、役所、病院、学校、職場、もう泣き寝入りしない！30分で読めるシリーズ.docx','22世紀の民主主義　選挙はアルゴリズムになり、政治家はネコになる (SB新書).docx','極上のおひとり死 (SB新書).docx','10秒で好かれるひとこと 嫌われるひとこと.docx','外資系コンサルの知的生産術～プロだけが知る「99の心得」～ (光文社新書).docx','職務経歴書はカタログだ！AI増補改訂版_ 書類選考突破率 90％の実績を持つ中途採用面接官が~える 「転職で勝てる職務経歴書」の極意！AI活用方法も大幅に加筆！ 電子書籍で学ぶ転職術.docx','新装版　不祥事 (講談社文庫).docx','新装版　銀行総務特命 (講談社文庫).docx','AI・データ分析プロジェクトのすべて[ビジネス力×技術力＝価値創出].docx','2025年版 新時代のIT経営戦略 知らなきゃ危ない! デジタル活用のリスクと対策 超実践入門_ これ一冊で、会社を守る具体的な対策がわかる！ (Volta Networks Publishing).docx','視点という教養（リベラルアーツ）　世界の見方が変わる７つの対話.docx','2025年版 新時代のIT経営戦略 AI活用で成果を最大化する方法_ 無駄な時間と労力に悩む人に向けた、AI活用の手引き (Volta Networks publishing).docx','小さいことにくよくよするな！.docx','AI 2041.docx','２０３５　１０年後のニッポン　ホリエモンの未来予測大全.docx','2025年度版ChatGPT「超」簡単誰でも画像生成入門_ 図解でわかる！GPT-4o対応の2025年最新版・画像生成AIの入門書 AI活用シリーズ.docx','超ヤバい経済学.docx','BCGカーボンニュートラル実践経営.docx','成瀬は都を駆け抜ける （「成瀬」シリーズ）_B0FQ4K3TW2.kfx.docx','生成AI×ビジネススキル_ 【仕事術】【効率化】【タイパ】【マーケティング】【セールス】【差別化】.docx','AIエージェント革命 「知能」を雇う時代へ_B0FCGQVV6V.kfx.docx','成瀬は信じた道をいく 「成瀬」シリーズ.docx','生成AIが起こすビジネス変革_ 未来のビジネスを予測する方法 Kindle生成AI解説本.docx','AI時代の質問力 プロンプトリテラシー 「問い」と「指示」が生成AIの可能性を最大限に引き出す.docx','AIが答えを出せない 問いの設定力.docx','AI×副業マネタイズ大全 ChatGPT・生成AI・TikTok・YouTube・アフィリエイト 初心者から月100万円ま~冊で全部わかるAI副業攻略本 月商1400万円＆300名のスクール経営者が教える「お金の稼ぎ方」.docx','人生を変える読書 人類三千年の叡智を力に変える.docx','破壊と創造の人事.docx','ChatGPT革命！6つの異次元プロンプト戦術で未来を切り開け！.docx','本質を見抜く「考え方」.docx','文章力が、最強の武器である。.docx','Cody_s Data Cleaning Techniques Using SAS, Third Edition.docx','文字コードの仕組みと歴史入門_ なぜ文字化けは起こるのか_ (ITプロ豆知識シリーズ).docx','ChatGPTで独学英語！_ TOEIC・英検の試験対策や面接・海外旅行の英会話などあらゆるシーンに対応できるAI英語学習_B0FFK7WP7X.kfx.docx','伝わる・揺さぶる！ 文章を書く (PHP新書).docx','ＢＣＧが読む経営の論点2022.docx','ChatGPTで新規ビジネス考案　_ 副業や新規事業で使えるプロンプト20選付.docx','売れないものを売る方法？ そんなものがほんとにあるなら教えてください！ (SB新書).docx','BRAIN DRIVEN　パフォーマンスが高まる脳の状態とは.docx','ChatGPT レシート自動記録ボット完全入門 _ コピペで15分作成！写真からエクセル（Googleスプレッドシート）に即保_B0FLJM77GB.kfx.docx','Designing Agentive Technology.docx','養老孟司の人生論 (PHP文庫).docx','Declutter Your Mind_ How to Stop Worrying, Relieve Anxiety, and Eliminate Negative Thinking.docx','ＣｈａｔＧＰＴ資本主義.docx','Claude-3.7 完全活用ガイド_ AI初心者でも今日から始められる業務効率化と収益化の実践バイブル.docx','ＣｈａｔＧＰＴ時代の文系ＡＩ人材になる―ＡＩを操る７つのチカラ.docx','ChatGPT×資料作成術_ プロンプトを活用し最速で上司が納得する資料を作成できる本。「生成AI~つけよう。【chatgptの頭の中】【API】【Excel】【英語】【副業】【チャットgpt】 ChatGPT仕事術.docx','ChatGPT Code Interpreter 実践解説 ４ Webデータベース.docx','Google開発「NotebookLM」決定版ガイド_ あなたの思考を自動で整理し、成果を最大化する_B0FHG6CJCC.kfx.docx','FULL POWER　科学が証明した自分を変える最強戦略.docx','FIRE生活への近道！正しい副業のはじめ方_ 凡人サラリーマンが貯金ゼロから3年で独立した方法 会社に頼らない生き方.docx','GAFAM＋テスラ 帝国の存亡 ビッグ・テック企業の未来はどうなるのか？.docx','ＧＡＦＡ　ｎｅｘｔ　ｓｔａｇｅ　ガーファ　ネクストステージ―四騎士＋Ｘの次なる支配戦略.docx','FACTFULNESS（ファクトフルネス）10の思い込みを乗り越え、データを基に世界を正しく見る習慣.docx','ExcelVBA［完全］入門.docx','LangChainとLangGraphによるRAG・AIエージェント［実践］入門 エンジニア選書.docx','Making Numbers Count_ The Art and Science of Communicating Numbers.docx','Lean Marketing_ More leads. More profit. Less marketing. (Lean Marketing Series).docx','Invent and Wander.docx','High Conflict_ Why We Get Trapped and How We Get Out.docx','GitLabに学ぶ 世界最先端のリモート組織のつくりかた ドキュメントの活用でオフィスなしでも最大の成果を出すグローバル企業のしくみ.docx','GO OUT (ゴーアウト) 飛び出す人だけが成功する時代.docx','Microsoft 365 Copilot AIで実現する仕事効率化.docx','ＬＩＦＥ　ＳＨＩＦＴ（ライフ・シフト）―１００年時代の人生戦略 LIFE SHIFT.docx','Microsoft Wordを開発した伝説のプログラマーが発見した「やりたいことの見つけ方」がすごい！.docx','LEARNING BPMN 2.0 A Practical Guide for Today’s Adult Learners_ An Introduction of Engineering Practices for Software Delivery Teams.docx','ＬＩＦＥＳＰＡＮ（ライフスパン）―老いなき世界.docx','ＬＩＦＥ　ＳＨＩＦＴ２―１００年時代の行動戦略 LIFE SHIFT.docx','OPEN TO THINK～最新研究が証明した 自分の小さな枠から抜け出す思考法.docx','ＯＲＩＧＩＮＡＬＳ　誰もが「人と違うこと」ができる時代.docx','NotebookLM最強勉強法_ ライバルはまだ知らないAIで差をつける最新勉強法_B0FSJTGF4B.kfx.docx','PLURALITY　対立を創造に変える、協働テクノロジーと民主主義の未来（サイボウズ式ブックス）_B0F17DLN9Y.kfx.docx','ＮＯＩＳＥ　下　組織はなぜ判断を誤るのか？.docx','ＮＯＩＳＥ　上　組織はなぜ判断を誤るのか？.docx','People Skills.docx','Markdownライティング入門　プレーンテキストで気楽に書こう！ (技術の泉シリーズ（NextPublishing）).docx','Power BI を最大限に活用する データモデリング.docx','Power Automateで学ぶローコード開発サバイバルガイド 技術の泉シリーズ.docx','PIXAR 〈ピクサー〉 世界一のアニメーション企業の今まで語られなかったお金の話.docx','PLG プロダクト・レッド・グロース「セールスがプロダクトを売る時代」から「プロダクトでプロダクトを売る時代」へ.docx','Pythonユニットテスト入門_ unittest、pytest、Coverage の基礎から応用まで Pythonエンジニア部.docx','Risk Up Front_ Managing Projects in a Complex World.docx','Q＆A　日本経済のニュースがわかる！　2021年版 (日本経済新聞出版).docx','Python FlaskによるWebアプリ開発入門 物体検知アプリ&機械学習APIの作り方.docx','The Values in Numbers.docx','Strategic Customer Service.docx','Strategy Safari_ A Guided Tour Through The Wilds of Strategic Management.docx','The Micro-Script Rules - 2nd Edition_ How to tell your story (and differentiate your brand) in a sentence...or less..docx','The Intelligent Sales AIを活用した最速・最良でクリエイティブな営業プロセス.docx','The Art of Laziness_ Overcome Procrastination & Improve Your Productivity.docx','The 1-Page Marketing Plan_ Get New Customers, Make More Money, And Stand Out From The Crowd (Lean Marketing Series).docx','SuperAgers スーパーエイジャー 老化は治療できる.docx','Work in Tech!（ワーク・イン・テック!） ユニコーン企業への招待.docx','Vibe Coding for Beginners with Python and ChatGPT.docx','THE YOROSHIKU GUIDE _ THE よろしく GUIDE_ Japanese Business Edition.docx','Webコピーライティングの新常識 ザ・マイクロコピー[第2版].docx','ＵＸグロースモデル　アフターデジタルを生き抜く実践方法論.docx','THIS IS MARKETING ディスイズマーケティング 市場を動かす.docx','TIME OFF 働き方に“生産性”と“創造性”を取り戻す戦略的休息術.docx','アウトプット思考 1の情報から10の答えを導き出すプロの技術.docx','アドバイスタロット_ タロット始めの一冊.docx','アキラとあきら　上 (集英社文庫).docx','あなたの収入が必ず増える!!　即断即決「脳」のつくり方.docx','アウトプットの精度を爆発的に高める「思考の整理」全技術.docx','イントゥ・ザ・ダーク 上 スター・ウォーズ ハイ・リパブリック.docx','いちばんやさしい機械学習プロジェクトの教本　人気講師が教える仕事にAIを導入する方法 「いちばんやさしい教本」シリーズ.docx','いつまで英語から逃げてるの？ 英語の多動力New Version 君の未来を変える英語のはなし.docx','アフターコロナのニュービジネス大全 新しい生活様式×世界15カ国の先進事例.docx','イノベーションの競争戦略―優れたイノベーターは０→１か？　横取りか？.docx','アフターデジタルセッションズ 最先端の33人が語る、世界標準のコンセンサス.docx','アナロジー思考.docx','アジャイルデータモデリング　組織にデータ分析を広めるためのテーブル設計ガイド (ＫＳ情報科学専門書).docx','オタク経済圏創世記　GAFAの次は2.5次元コミュニティが世界の主役になる件.docx','うまくいっている会社の非常識な儲け方.docx','エンベデッド・ファイナンスの衝撃―すべての企業は金融サービス企業になる.docx','エンジニアを説明上手にする本 相手に応じた技術情報や知識の伝え方.docx','エンジニアのためのWord再入門講座 新版 美しくメンテナンス性の高い開発ドキュメントの作り方.docx','エフォートレス思考 努力を最小化して成果を最大化する.docx','コンサル一〇〇年史 (ディスカヴァー・レボリューションズ).docx','キミは、「怒る」以外の方法を知らないだけなんだ.docx','コンサルを超える　問題解決と価値創造の全技法.docx','コミュニティシフト_ すべてがコミュニティ化する時代.docx','コマンドラインの黒い画面が怖いんです。 新人エンジニアのためのコマンドが使いこなせる本.docx','かばん屋の相続.docx','ガウス過程と機械学習 (機械学習プロフェッショナルシリーズ).docx','サピエンス全史　上下合本版　文明の構造と人類の幸福.docx','サイゼリヤ元社長が教える 年間客数２億人の経営術.docx','サイコパスに学ぶ成功法則.docx','コンセプチュアル思考 物事の本質を見極め、解釈し、獲得する.docx','シン・ロジカルシンキング.docx','シン・ニホン AI×データ時代における日本の再生と人材育成 (NewsPicksパブリッシング).docx','システム導入のためのデータ移行ガイドブック―コンサルタントが現場で体得したデータ移行のコツ (NextPublishing).docx','シンプルに結果を出す人の　５Ｗ１Ｈ思考.docx','システムを作らせる技術　エンジニアではないあなたへ (日本経済新聞出版).docx','スモールビジネスの時代がやってくる.docx','すべては「前向き質問」でうまくいく 質問思考の技術 増補改訂版.docx','すぐやる！　「行動力」を高める“科学的な”方法.docx','スター・ウォーズ アソーカ 下.docx','ストーリーで学ぶ新規ビジネス創造戦略_ 新しい市場を開拓する！ビジネスアイデアを形に~スケースを通じて、競争に勝つための戦略を実践的に学ぶ 仕事で実績を作るために役立つ本.docx','スター・ウォーズ ハイ・リパブリック2 イントゥ・ザ・ダーク 下.docx','スター・ウォーズ アソーカ 上.docx','ゼロ・トゥ・ワン　君はゼロから何を生み出せるか.docx','ソフトウェア開発現場の「失敗」集めてみた。 42の失敗事例で学ぶチーム開発のうまい進めかた.docx','ダイアグラム思考 次世代型リーダーは図解でチームを動かす.docx','その「一言」が子どもの脳をダメにする (SB新書).docx','ゼロからわかる生成AI駆動開発入門_ ハンズオンで学ぶ Bolt.new, ChatGPT, Replit Agent, Cursor, GEAR.indigo, Feloなど最新ツールを活用した最先端の開発アプローチ.docx','だから僕たちは、組織を変えていける ――やる気に満ちた「やさしいチーム」のつくりかた【ビジネス書グランプリ2023「マネジメント部門賞」受賞！】.docx','ダブルハーベスト――勝ち続ける仕組みをつくるＡＩ時代の戦略デザイン.docx','タロットと遊んで未来をつくる方法_ リアル宝探しゲーム編 タロットで未来創造シリーズ (占わない師の道具箱).docx','タロット心理学の入門書_ タロット占いをセラピーレベルに引き上げるスピリチュアルと学問の交差点.docx','タロットカード 全78種 解説ハンドブック_ 意味・種類・解釈・占い・占星術.docx','タピオカ屋はどこへいったのか？　商売の始め方と儲け方がわかるビジネスのカラクリ.docx','ツキの最強法則【CD無し】.docx','チャートで考えればうまくいく 一生役立つ「構造化思考」養成講座 (セブンチャートテンプレート特典付き).docx','チーム・ジャーニー 逆境を越える、変化に強いチームをつくりあげるまで.docx','デジノグラフィ インサイト発見のためのビッグデータ分析.docx','データ分析・AIを実務に活かす データドリブン思考.docx','できる大人のモノの言い方大全　LEVEL2.docx','デジタルトランスフォーメーションに伴う科学技術・イノベーションの変容（—The Beyond Disciplines Collection—）.docx','データ戦略と法律　攻めのビジネスQ＆A　改訂版.docx','データ可視化の基本が全部わかる本 収集・変換からビジュアライゼーション・データ分析支援まで.docx','チームレジリエンス　困難と不確実性に強いチームのつくり方.docx','デジタル時代のカスタマーサービス戦略.docx','トランスフォーメーション思考 未来に没入して個人と組織を変革する.docx','トリーズ(TRIZ)の発明原理40 あらゆる問題解決に使える[科学的]思考支援ツール トリーズ（TRIZ）の発明原理４０.docx','ドラマ「半沢直樹」原作　ロスジェネの逆襲.docx','トーキング・トゥ・ストレンジャーズ～「よく知らない人」について私たちが知っておくべきこと～.docx','テクノ・リバタリアン　世界を変える唯一の思想 (文春新書).docx','デジタル未来にどう変わるか？　AIと共存する個人と組織.docx','パッション・パラドックス.docx','ハーバード流「聞く」技術 (角川新書).docx','ノーサイド・ゲーム.docx','なぜあの人の解決策はいつもうまくいくのか？―小さな力で大きく動かす！システム思考の上手な使い方.docx','はじめての課長の教科書 第3版.docx','バイブコーディング完全入門_ 非エンジニアがAIを活用してプログラムを作り 　自動化＆効率化を実現するため_B0FBKTTN55.kfx.docx','ニュー・エリートの時代　ポストコロナ「３つの二極化」を乗り越える (角川書店単行本).docx','なぜ、倒産寸前の水道屋がタピオカブームを仕掛け、アパレルでも売れたのか？.docx','なぜ、サボる人ほど成果があがるのか？　仕事の速い人になる時間術101.docx','なぜ、あの人は仕事ができるのか？.docx','どんなビジネスを選べばいいかわからない君へ.docx','なぜか結果を出す人が勉強以前にやっていること.docx','ドラッカーと会計の話をしよう (中経の文庫).docx','プロジェクトのトラブル解決大全　小さな問題から大炎上まで使える「プロの火消し術86」.docx','プロジェクトマネジメントの基本が全部わかる本 交渉・タスクマネジメント・計画立案から見積り・契約・要件定義・設計・テスト・保守改善まで.docx','ファンタスティック・ビーストと魔法使いの旅　〈映画オリジナル脚本版〉 (ファンタスティック・ビースト (Fantastic Beasts)).docx','ブレイクスルー ひらめきはロジックから生まれる.docx','ひとつ上の思考力.docx','フューチャリストの「自分の未来」を変える授業.docx','ビジネスの構築から最新技術までを網羅　AIの教科書.docx','ファンタスティック・ビーストとダンブルドアの秘密　映画オリジナル脚本版_ 「ファンタ~ト」映画オリジナル脚本版シリーズ　第3巻 (ファンタスティック・ビースト (Fantastic Beasts)).docx','ビジネスを育てる 新版 いつの時代も変わらない起業と経営の本質.docx','ビジネスに魔法をかける　生成AI導入大全.docx','ひとりビジネスの教科書 Premium 自宅起業でお金と自由を手に入れて成功する方法.docx','ピープルアナリティクスの教科書.docx','ビジネスモデル2.0図鑑 (中経出版).docx','マッキンゼーのエリートはノートに何を書いているのか　トップコンサルタントの考える技術・書く技術.docx','まず、ちゃんと聴く。　コミュニケーションの質が変わる｢聴く｣と｢伝える｣の黄金比.docx','マンガでわかるChatGPT4-o超活用術【図解】_ 初心者でもマネするだけ20分でGPT4-oの活用法がわ~】【音声対話】【ロープレ】【英会話】【学習コーチ】 ChatGPT副業プロンプト入門シリーズ.docx','マーケティングリサーチとデータ分析の基本.docx','マッキンゼーで叩き込まれた 超速フレームワーク―――仕事のスピードと質を上げる最強ツール (三笠書房　電子書籍).docx','マッキンゼーが解き明かす 生き残るためのDX (日本経済新聞出版).docx','ボイステック革命 　GAFAも狙う新市場争奪戦 (日本経済新聞出版).docx','ホモ・デウス　上下合本版　テクノロジーとサピエンスの未来.docx','マッキンゼー　ネクスト・ノーマル―アフターコロナの勝者の条件.docx','マジビジプロ 新人コンサルタントが最初に学ぶ 厳選フレームワーク20 MAJIBIJI pro［図解］問題解決に強くなる！.docx','ヤバい集中力　1日ブッ通しでアタマが冴えわたる神ライフハック45.docx','ヤバい経営学―世界のビジネスで行われている不都合な真実.docx','もっと早く、もっと楽しく、仕事の成果をあげる法 知恵がどんどん湧く「戦略的思考力」を身につけよ.docx','メタ思考トレーニング 発想力が飛躍的にアップする34問 PHPビジネス新書.docx','ムダがなくなり、すべてがうまくいく 本当の時間術.docx','マンガでわかるAIエージェント～OpenAI「operator（オペレーター）」&「OpenManus（オープンマヌ~わりに働く秘書になる。無料の環境構築から実用例まで徹底解説！ マンガでわかるシリーズ.docx','マンガで読める マッキンゼー流「問題解決」がわかる本.docx','リサーチ・ドリブン・イノベーション 「問い」を起点にアイデアを探究する.docx','ヤバい経済学〔増補改訂版〕―悪ガキ教授が世の裏側を探検する.docx','リーダーは話し方が９割.docx','ユニクロの仕組み化.docx','ようこそ、わが家へ.docx','リーダーのように組織で働く.docx','ロジカル・シンキング Best solution.docx','レヴィット　ミクロ経済学　基礎編.docx','ロジカル・シンキング練習帳―論理的な考え方と書き方の基本を学ぶ５１問.docx','レヴィット　ミクロ経済学　発展編.docx','ロジカル・ライティング―論理的にわかりやすく書くスキル BEST SOLUTION―LOGICAL COMMUNICATION SKILL TRAINING.docx','医療現場の行動経済学―すれ違う医者と患者.docx','一生仕事で困らない企画のメモ技(テク)―――売れる企画を“仕組み”で生み出すメモの技術.docx','一度読んだら絶対に忘れない英文法の教科書.docx','一瞬で人の心を操る「売れる」セールスライティング.docx','ロジカルタロット　占いを超えた新しいリーディング・メソッド.docx','億を稼ぐ言葉の技術29_ 〜心理学×コピーライティングの教本〜.docx','仮説とデータをつなぐ思考法　DATA INFORMED.docx','仮説行動――マップ・ループ・リープで学びを最大化し、大胆な未来を実現する.docx','映画ノベライズ 【推しの子】－The Final Act－ (集英社オレンジ文庫).docx','右脳思考を鍛える―「観・感・勘」を実践！　究極のアイデアのつくり方.docx','億を売る『LP理論』_ 〜最高の見込み客リスト獲得ランディングページ教本〜.docx','因果推論 ―基礎から機械学習・時系列解析・因果探索を用いた意思決定のアプローチ―.docx','一度読んだら絶対に忘れない日本史の教科書　公立高校教師YouTuberが書いた.docx','一度読んだら絶対に忘れない世界史人物事典　公立高校教師YouTuberが書いた.docx','一度読んだら絶対に忘れない世界史の教科書　公立高校教師YouTuberが書いた.docx','一度読んだら絶対に忘れない化学の教科書.docx','改訂版　タロットカードの読み方 78枚のカードの2000を超える解説　第１巻_ タロットカードの意味 (LearnCreate).docx','改訂版　タロットカードの読み方 78枚のカードの2000を超える解説　第２巻_ タロットカードの意味 (LearnCreate).docx','会社の問題発見、課題設定、問題解決.docx','箇条書きの習慣.docx','花咲舞が黙ってない (中公文庫).docx','果つる底なき (講談社文庫).docx','架空通貨 (講談社文庫).docx','下町ロケット　ヤタガラス.docx','感動させて→行動させる エモいプレゼン.docx','管理職はいらない　AI時代のシン・キャリア (SB新書).docx','株価暴落.docx','外資系コンサルのPowerPoint資料作成140のテクニック_ スライドをわかりやすく、速くつくる理論と技術.docx','外資系データサイエンティストの知的生産術―どこへ行っても通用する人になる超基本５０.docx','外資系コンサルのビジネス文書作成術―ロジカルシンキングと文章術によるＷｏｒｄ文書の作り方.docx','機械学習を解釈する技術〜予測力と説明力を両立する実践テクニック.docx','逆襲のビジネス教室.docx','企画立案からシステム開発まで　本当に使えるDXプロジェクトの教科書.docx','逆境を「アイデア」に変える企画術　～崖っぷちからV字回復するための40の公式～.docx','観察力の鍛え方　一流のクリエイターは世界をどう見ているのか (SB新書).docx','現場で使える！pandasデータ前処理入門 機械学習・データサイエンスで役立つ前処理手法.docx','現場で使える! AI活用入門 技術の泉シリーズ.docx','結果を出し続ける人が行動する前に考えていること 無理が勝手に無理でなくなる仕組みの作り方.docx','経済評論家の父から息子への手紙 お金と人生と幸せについて.docx','結果を引き寄せる　完全版　YouTube TikTokビジネス活用術.docx','銀行狐 (講談社文庫).docx','競争戦略論（第２版）.docx','教養としての生成AI (幻冬舎新書).docx','今日がもっと楽しくなる行動最適化大全　ベストタイムにベストルーティンで常に「最高の1日」を作り出す.docx','顧客体験マーケティング 顧客の変化を読み解いて「売れる」を再現する（Web担選書）.docx','最強の生き方.docx','御社の新規事業はなぜ失敗するのか？～企業発イノベーションの科学～ (光文社新書).docx','効果検証入門〜正しい比較のための因果推論／計量経済学の基礎.docx','高業績を上げ続ける「三点思考営業」 安定収益型の最幸最強営業を目指す.docx','言語化の魔力　言葉にすれば「悩み」は消える (幻冬舎単行本).docx','今すぐ始められる「ショート動画編集」の教科書　この一冊でOK！未経験から最速で月10万円稼ぐ方法.docx','現役LPO会社社長から学ぶ コンバージョンを獲る ランディングページ.docx','交通事故事件対応のための保険の基本と実務.docx','言語化力　言葉にできれば人生は変わる.docx','現場で活用するためのＡＩエージェント実践入門 (ＫＳ情報科学専門書)_B0FM3QJ2DP.kfx.docx','思考の質を高める 構造を読み解く力.docx','子どもの幸せを一番に考えるのをやめなさい (SB新書).docx','使ってわかったAWSのAI ～まるごと試せば視界は良好 さあPythonではじめよう！.docx','仕事と勉強を両立させる時間術.docx','最高の打ち手が見つかるマーケティングの実践ガイド 3つのマップで戦略に沿った施策を実行する.docx','仕事は楽しいかね？ (きこ書房).docx','仕事選びのアートとサイエンス～不確実な時代の天職探し　改訂『天職は寝て待て』～ (光文社新書).docx','山崎元のライフマネジメント 幸せな人生のための基本戦略.docx','仕事の9割は数学思考でうまくいく(あさ出版電子書籍).docx','最短で最高の結果を出す副業バイブル.docx','最終退行 (小学館文庫).docx','最高の脳で働く方法　Your Brain at Work.docx','失敗の科学　失敗から学習する組織、学習できない組織.docx','七つの会議 (日本経済新聞出版).docx','失敗学実践講義　文庫増補版 (講談社文庫).docx','持続可能なチームのつくり方 幸福と成果が連動する.docx','紫式部 (やさしく読める ビジュアル伝記).docx','自分の時間―――１日２４時間でどう生きるか (三笠書房　電子書籍).docx','自分とチームの生産性を最大化する 最新「仕組み」仕事術.docx','次のテクノロジーで世界はどう変わるのか (講談社現代新書).docx','死はこわくない (文春文庫).docx','施策デザインのための機械学習入門〜データ分析技術のビジネス活用における正しい考え方.docx','視点転換転職術_ 転職活動を勝ち進む5つのノウハウ 電子書籍で学ぶ転職 (KBLT出版)_B0FLQ7CDGP.kfx.docx','仕事力を爆上げする「図解思考」 (三笠書房　電子書籍).docx','社内外に眠るデータをどう生かすか ―データに意味を見出す着眼点― (養成講座シリーズ).docx','実践Claude Code入門―現場で活用するためのAIコーディングの思考法 エンジニア選書_B0G4NVX1NF.kfx.docx','実践版ＧＲＩＴ　やり抜く力を手に入れる.docx','実践Androidアプリシステムテスト　実機でテスト！チームで挑む品質向上入門 技術の泉シリーズ.docx','自分を変える話し方.docx','実践　医療現場の行動経済学―すれ違いの解消法.docx','実践型クリティカルシンキング 特装版.docx','情弱の解剖学_ 人はなぜ弱者化し、どうすれば抜け出せるのか？ (ビジネス書).docx','小さくはじめよう ー自分らしい事業を手づくりできる「マイクロ起業」メソッド.docx','小さく分ければうまくいく_ 「待ち時間」を宝に変える 身近な仕事改革のヒント 制約を愉しむシリーズ (たくらみKnowledge).docx','上流思考――「問題が起こる前」に解決する新しい問題解決の思考法.docx','集中力がすべてを解決する　精神科医が教える「ゾーン」に入る方法.docx','勝者の科学 一流になる人とチームの法則.docx','小さく分けて考える　「悩む時間」と「無駄な頑張り」を80％減らす分解思考.docx','図解　無印良品は、仕組みが９割　仕事はシンプルにやりなさい (角川書店単行本).docx','図解 オンライン研修入門.docx','人を動かすマーケティングの新戦略「行動デザイン」の教科書.docx','新世代のビジネスはスマホの中から生まれる ショートムービー時代のSNSマーケティング.docx','新しい経営学.docx','人生を変える「質問力」の教え.docx','人工知能は人間を超えるか (角川ＥＰＵＢ選書).docx','新規事業のセオリー_ ５つのフェーズ、４１のステップで進める実践本 (ビジネス 経済).docx','身体を壊す健康法 年間500本以上読破の論文オタクの東大医学博士＆現役医師が、世界中から有益な情報を見つけて解き明かす。.docx','新装版　問題解決のためのデータ分析.docx','情報吸収力を高めるキーワード読書術.docx','情報を活用して、思考と行動を進化させる.docx','世界はなぜ地獄になるのか（小学館新書）.docx','世界の起業家が学んでいるＭＢＡ経営理論の必読書５０冊を１冊にまとめてみた.docx','数学ガールの誕生　理想の数学対話を求めて.docx','推しエコノミー　「仮想一等地」が変えるエンタメの未来.docx','世界3万人のハイパフォーマー分析でわかった 成功し続ける人の6つの習慣.docx','数学的思考トレーニング 問題解決力が飛躍的にアップする48問 (PHPビジネス新書).docx','数学嫌いな人のための数学（新装版）.docx','図解でサクッと理解！週末3日でゲームをリリース！ITエンジニアのためのAIエージェント副業術_ この1冊だけでゲームが完成するープロンプト集などの2大特典付き.docx','世界一流エンジニアの思考法 (文春e-book).docx','成功をめざす人に知っておいてほしいこと 新版.docx','世界史とつなげて学べ 超日本史　日本人を覚醒させる教科書が教えない歴史.docx','世界最高峰の研究者たちが予測する未来 (SB新書).docx','世界はラテン語でできている (SB新書).docx','世界一やさしい「やりたいこと」の見つけ方　人生のモヤモヤから解放される自己理解メソッド.docx','生成ＡＩ時代の「超」仕事術大全.docx','生成AIで世界はこう変わる (SB新書).docx','成功率を高める新規事業のつくり方 顧客・機能・技術の3軸で新しいビジネスアイデアを創出する方法.docx','世界最先端の研究が導き出した、「すぐやる」超習慣.docx','正規表現辞典 改訂新版.docx','生成AI導入の教科書.docx','成瀬は天下を取りにいく（新潮文庫） （「成瀬」シリーズ）_B0F4259M1G.kfx.docx','戦略論と科学思考の教科書_ チームで創るデータ駆動イノベーション.docx','戦略的いい人 残念ないい人の考え方.docx','戦わずして売る技術　クリック１つで市場を生み出す最強のWEBマーケティング術 (幻冬舎単行本)_B0FMQ1Q69F.kfx.docx','戦略質問―短時間だからこそ優れた打ち手がひらめく.docx','先輩、これからボクたちは、どうやって儲けていけばいいんですか？.docx','生命保険の不都合な真実 (光文社新書).docx','責任あるＡＩ―「ＡＩ倫理」戦略ハンドブック.docx','相手を完全に信じ込ませる禁断の心理話術　エニアプロファイル.docx','走りながら考える 新規事業の教科書.docx','相手に刺さる話し方.docx','早く読めて、忘れない、思考力が深まる　「紙１枚！」読書法.docx','全米ナンバーワンビジネススクールで教える起業家の思考と実践術―あなたも世界を変える起業家になる.docx','創造はシステムである　「失敗学」から「創造学」へ (角川oneテーマ21).docx','戦略読書日記_本質を抉りだす思考のセンス_.docx','全社デジタル戦略 失敗の本質_B0FF3W82VB.kfx.docx','大学4年間のデータサイエンスが10時間でざっと学べる (角川文庫).docx','大学4年間のマーケティングが10時間でざっと学べる (角川文庫).docx','多様性の科学　画一的で凋落する組織、複数の視点で問題を解決する組織.docx','対峙力ーー誰にでも堂々と振る舞えるコミュニケーション術.docx','大下流国家～「オワコン日本」の現在地～ (光文社新書).docx','地頭力を鍛える.docx','地球の未来のため僕が決断したこと　気候大災害は防げる.docx','誰も教えてくれなかったアジャイル開発.docx','短い言葉を武器にする.docx','知識ゼロから学ぶソフトウェアテスト 第3版 アジャイル・AI時代の必携教科書.docx','大学4年間の社会心理学が10時間でざっと学べる.docx','大学4年間の金融学が10時間でざっと学べる (角川文庫).docx','虫とゴリラ.docx','直感力を高める　数学脳のつくりかた (河出文庫).docx','超訳　孫子の兵法　「最後に勝つ人」の絶対ルール.docx','超リテラシー大全.docx','池上彰の未来予測　After 2040.docx','東京23区×格差と階級 (中公新書ラクレ).docx','転職2.0　日本人のキャリアの新・ルール.docx','伝える準備.docx','伝え方――伝えたいことを、伝えてはいけない。.docx','伝わる　イラスト思考.docx','日本人「総奴隷化」計画　１９８５ー２０２９　アナタの財布を狙う「国家の野望」.docx','届く！刺さる！！売れる！！！　キャッチコピーの極意.docx','読書する人だけがたどり着ける場所 (SB新書).docx','独学の地図.docx','独習Python.docx','入門『地頭力を鍛える』　３２のキーワードで学ぶ思考法.docx','任せるコツ.docx','認知症にならない ストレスマネジメント　医師が実践する 脳ダメージをはねのける方法.docx','入門 Dify_ 1時間で学ぶ基礎+エージェント・RAG作成 1時間で学ぶLLMツール.docx','日本人の9割が知らない遺伝の真実 (SB新書).docx','入社１年目から差がつく　ロジカル・シンキング練習帳.docx','犯罪心理学者が教える子どもを呪う言葉・救う言葉 (SB新書).docx','描くだけで毎日がハッピーになる　ふだん使いのマインドマップ.docx','不完全主義　限りある人生を上手に過ごす方法_B0FCDYRN3M.kfx.docx','武器としての図で考える習慣―「抽象化思考」のレッスン.docx','半沢直樹　１　オレたちバブル入行組 (講談社文庫).docx','売れる言いかえ大全.docx','未来を実装する――テクノロジーで社会を変革する４つの原則.docx','弁護士費用特約を活用した物損交通事故の実務.docx','暮らしも心も調う大人の断捨離手帖.docx','米中先進事例に学ぶ マーケティングDX.docx','片田舎のおっさん、剣聖になる　1　～ただの田舎の剣術師範だったのに、大成した弟子たちが俺を放ってくれない件～ (デジタル版SQEXノベル).docx','文系ＡＩ人材になる―統計・プログラム知識は不要.docx','副業をはじめたいんですけど、税金ってどうしたらいいですか？.docx','問題解決のジレンマ―イグノランスマネジメント：無知の力.docx','目的ドリブンの思考法.docx','無駄な仕事が全部消える超効率ハック.docx','問題解決 ― あらゆる課題を突破する ビジネスパーソン必須の仕事術.docx','民王 (角川文庫).docx','儲ける仕組みをつくるフレームワークの教科書.docx','良いFAQの書き方──ユーザーの「わからない」を解決するための文章術 WEB+DB PRESS plus.docx','劣化するオッサン社会の処方箋～なぜ一流は三流に牛耳られるのか～ (光文社新書).docx','陸王 (集英社文庫).docx','利益を最大化する　価格決定戦略.docx','熔ける 再び　そして会社も失った (幻冬舎単行本).docx','問題解決ファシリテーター―「ファシリテーション能力」養成講座.docx','予測マシンの世紀　AIが駆動する新たな経済.docx','話を聞かない男、地図が読めない女.docx','話し方で老害になる人尊敬される人 若者との正しい話し方&距離感 正解・不正解.docx']
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



