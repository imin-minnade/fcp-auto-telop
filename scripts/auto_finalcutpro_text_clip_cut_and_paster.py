"""Final Cut Pro のテキストクリップを自動分割し、セリフを貼り付けるスクリプト。

CSV ファイルからセリフを読み込み、FCP のタイムライン上で
テキストクリップを自動的に分割・貼り付けします。

使い方:
1. csv_input/ にシナリオ CSV を配置（列: 実行, キャラクター, セリフ）
2. CSV_FILE を対象ファイル名に変更
3. TARGET_CHARACTER を対象キャラクター名に変更
4. INPUT_X, INPUT_Y を calibrate_coordinates.py で取得した値に変更
5. FCP でテキストクリップが選択された状態にする
6. `python scripts/auto_finalcutpro_text_clip_cut_and_paster.py` を実行
"""

import os
import sys
import platform
import pyautogui
import time
import pandas as pd
import pyperclip

# ===================== 設定 =====================
# シナリオ CSV ファイル（csv_input/ ディレクトリに配置）
CSV_FILE = "csv_input/sample.csv"

# 対象キャラクター名（CSV の「キャラクター」列と一致するもの）
TARGET_CHARACTER = "ナレーター"

# テキスト入力欄の座標（calibrate_coordinates.py で取得）
INPUT_X, INPUT_Y = 3457, 190

# 合成音声モード: True = 合成音声（シナリオからセリフ取得）
#                 False = 生声（Whisper 文字起こしからセリフ取得）
USE_AI_VOICE = True
# ===================== 設定ここまで =====================


def load_voices_from_csv(csv_path: str, target: str) -> list[str]:
    """CSV ファイルからセリフを読み込む。

    Args:
        csv_path: CSV ファイルのパス
        target: 対象キャラクター名

    Returns:
        セリフのリスト
    """
    if not os.path.exists(csv_path):
        print(f"❌ CSV ファイルが見つかりません: {csv_path}")
        print(f"   csv_input/ ディレクトリにファイルを配置してください。")
        sys.exit(1)

    df = pd.read_csv(csv_path, encoding="utf-8", header=0)

    # 必要な列の存在チェック
    required_cols = ["実行", "キャラクター", "セリフ"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        print(f"❌ CSV に必要な列がありません: {missing}")
        print(f"   必要な列: {required_cols}")
        sys.exit(1)

    # 実行列を数値型に変換
    df["実行"] = pd.to_numeric(df["実行"], errors="coerce")

    # フィルタリング: 実行=1 かつ 対象キャラクター かつ セリフが空でない
    mask = (df["実行"] == 1) & (df["キャラクター"] == target) & (df["セリフ"].notna())
    voices = df.loc[mask, "セリフ"].tolist()

    if not voices:
        print(f"❌ 対象のセリフが見つかりません（キャラクター: {target}）")
        sys.exit(1)

    return voices


def main():
    # CSV からセリフを読み込み
    voice_list = load_voices_from_csv(CSV_FILE, TARGET_CHARACTER)

    print(f"📄 CSV ファイル: {CSV_FILE}")
    print(f"🎭 キャラクター: {TARGET_CHARACTER}")
    print(f"📝 セリフの数: {len(voice_list)}")
    print()

    # 確認表示
    for i, v in enumerate(voice_list):
        print(f"  {i + 1}. {v}")
    print()

    # 後ろから入力するため逆順にする
    voice_list = voice_list[::-1]

    print("準備")
    print("テキストクリップを分割しておく。テキストフィールドが見える状態にしておきます。")
    print("⏱ 5秒後に開始します。Final Cut Pro をアクティブにしておいてください！")
    time.sleep(5)

    # テキストクリップを分割
    if USE_AI_VOICE:
        cut_count = len(voice_list) - 2
    else:
        cut_count = len(voice_list) * 2 - 1

    for _ in range(cut_count):
        # 次のクリップに移動
        pyautogui.keyDown("command")
        pyautogui.press("right")
        pyautogui.keyUp("command")

        # ボイス.mp3 の接合箇所に移動
        time.sleep(0.5)
        pyautogui.press("down")

        time.sleep(0.5)
        # テキストクリップを分割（Command + B）
        pyautogui.keyDown("command")
        pyautogui.press("b")
        pyautogui.keyUp("command")

        time.sleep(0.5)

    # 最後のテキストクリップに移動する
    if USE_AI_VOICE:
        pyautogui.keyDown("command")
        pyautogui.press("right")
        pyautogui.keyUp("command")
        time.sleep(3)

    # ループでセリフを入力する
    for i, voice in enumerate(voice_list):
        print(f"  [{i + 1}/{len(voice_list)}] {voice}")

        # クリップボードにコピー
        time.sleep(0.5)
        pyperclip.copy(voice)

        # テキストフィールドに移動
        time.sleep(0.5)
        if i == 0:
            # 最初はテキストエリアをクリック
            pyautogui.click(INPUT_X, INPUT_Y)
            time.sleep(0.5)
        else:
            # 2回目以降は Tab で移動
            pyautogui.press("tab")

        # 貼り付け（Command + V）
        time.sleep(0.5)
        pyautogui.keyDown("command")
        pyautogui.press("v")
        pyautogui.keyUp("command")

        # 前のクリップへ移動（Command + Left）
        time.sleep(0.5)
        pyautogui.keyDown("command")
        pyautogui.press("left")
        pyautogui.keyUp("command")

        # 生声モードの場合は、もう1つ前のクリップに移動
        if not USE_AI_VOICE:
            time.sleep(0.5)
            pyautogui.keyDown("command")
            pyautogui.press("left")
            pyautogui.keyUp("command")

    print("✅ すべてのテロップを入力しました")


if __name__ == "__main__":
    main()
