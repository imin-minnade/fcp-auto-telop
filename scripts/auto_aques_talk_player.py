"""AquesTalk Player にセリフを自動入力して読み上げるスクリプト。

CSV ファイルからセリフを読み込み、AquesTalk Player に
自動でテキストを入力して音声を生成します。

使い方:
1. csv_input/ にシナリオ CSV を配置（列: 実行, キャラクター, セリフ）
2. CSV_FILE, TARGET_CHARACTER を変更
3. INPUT_X, INPUT_Y, BUTTON_X, BUTTON_Y を get_mouse_positions.py で取得した値に変更
4. AquesTalk Player を開いた状態で実行
"""

import platform
import pyautogui
import time
import random
import pandas as pd
import pyperclip

# ===================== 設定 =====================
# シナリオ CSV ファイル（csv_input/ ディレクトリに配置）
CSV_FILE = "csv_input/sample.csv"

# 読み上げ対象のキャラクター名
TARGET_CHARACTER = "魔理沙"

# AquesTalk Player の入力欄座標（get_mouse_positions.py で取得）
INPUT_X, INPUT_Y = 36, 97

# 再生ボタンの座標
BUTTON_X, BUTTON_Y = 349, 194

# 開始前のカウントダウン（秒）
SLEEP_COUNTDOWN = 5
# ===================== 設定ここまで =====================

# 読みの修正リスト（AquesTalk が正しく読めない単語を修正）
REPLACE_LIST = [
    ["ChatGPT", "チャットジーピーティー"],
    ["Python", "パイソン"],
    ["YouTube", "ユーチューブ"],
    ["VOICEVOX", "ボイスボックス"],
    ["Final Cut Pro", "ファイナルカットプロ"],
    ["GitHub", "ギットハブ"],
    ["Whisper", "ウィスパー"],
    ["一行", "いちぎょう"],
    ["他人", "たにん"],
    ["後で", "あとで"],
    # 必要に応じて追加してください
]


def wait_random_interval():
    """読み上げ完了を待つランダムインターバル。"""
    wait_time = random.randint(1, 2)
    print(f"⏳ {wait_time}秒待機...")
    time.sleep(wait_time)


# CSV 読み込み
df = pd.read_csv(CSV_FILE, encoding="utf-8", header=0)

print(f"📄 CSV ファイル: {CSV_FILE}")
print(f"🎭 キャラクター: {TARGET_CHARACTER}")
print(f"   CSV の列名: {df.columns.tolist()}")

# 実行列を数値型に変換
df["実行"] = pd.to_numeric(df["実行"], errors="coerce")

print(f"⏱ {SLEEP_COUNTDOWN}秒後に開始します。AquesTalk Player をアクティブにしておいてください！")
time.sleep(SLEEP_COUNTDOWN)

# OS によって修飾キーを切り替え（ループ外で1回だけ取得）
modifier_key = "command" if platform.system() == "Darwin" else "ctrl"

count = 0
for _, row in df.iterrows():
    # 実行条件のチェック
    if pd.isna(row["実行"]) or row["実行"] != 1:
        continue

    # 対象キャラクターのチェック
    if TARGET_CHARACTER != row["キャラクター"]:
        continue

    voice = row["セリフ"]

    if pd.isna(voice):
        continue

    # 読み方を修正
    for replace_item in REPLACE_LIST:
        if replace_item[0] in voice:
            voice = voice.replace(replace_item[0], replace_item[1])
            print(f"  読み修正: {replace_item[0]} → {replace_item[1]}")

    count += 1
    print(f"🖊 [{count}] {voice}")

    # 入力欄をクリックしてフォーカス
    pyautogui.click(INPUT_X, INPUT_Y)
    time.sleep(0.5)

    # クリップボードにコピー
    pyperclip.copy(voice)

    # テキスト全選択
    pyautogui.keyDown(modifier_key)
    pyautogui.press("a")
    pyautogui.keyUp(modifier_key)

    # テキスト貼り付け
    pyautogui.keyDown(modifier_key)
    pyautogui.press("v")
    pyautogui.keyUp(modifier_key)

    # 再生ボタンをクリック
    time.sleep(1)
    pyautogui.click(BUTTON_X, BUTTON_Y)

    wait_random_interval()

print(f"✅ すべてのセリフを送信しました（{count}件）")
