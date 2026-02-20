import os
import sys
import time
from decimal import Decimal, ROUND_HALF_UP
import pyautogui
import pyperclip

# ====== 環境依存パラメータ（あなたの環境に合わせて調整） ======

# VTT ファイル（vtt_input/ ディレクトリに配置）
VTT_FILE = "vtt_input/音声.vtt"

# テキストインスペクタのテキスト入力フィールドの位置（画面座標）
INPUT_X, INPUT_Y = 3457, 190

# タイムラインのフレームレート
FPS = 25  # プロジェクトに合わせて変更

# セリフ間隔の最小ギャップ（秒）。これ以下なら開始時刻を押し出す。
MIN_GAP_SEC = 0.1

# キー操作・クリック後のウェイト（秒）
SLEEP_SHORT = 0.5
# クリップ移動・FCP描画完了を待つウェイト（秒）
SLEEP_CLIP_MOVE = 1
# 開始前のカウントダウン（秒）
SLEEP_COUNTDOWN = 5

# ====== 設定ここまで ======


# =====================================================
# VTT パーサー
# =====================================================

def parse_vtt_from_file(vtt_path: str):
    """VTT ファイルをパースして、[{start, end, text}, ...] を返す。"""
    if not os.path.exists(vtt_path):
        print(f"❌ VTT ファイルが見つかりません: {vtt_path}")
        print(f"   vtt_input/ ディレクトリにファイルを配置してください。")
        sys.exit(1)

    with open(vtt_path, encoding="utf-8") as f:
        text = f.read()

    blocks = text.strip().split("\n\n")
    cues = []
    for block in blocks:
        b = block.strip()
        if not b:
            continue
        if b.startswith("WEBVTT"):
            continue

        lines = [l for l in b.splitlines() if l.strip() != ""]
        if not lines:
            continue

        time_line_index = None
        for idx, line in enumerate(lines):
            if "-->" in line:
                time_line_index = idx
                break
        if time_line_index is None:
            continue

        time_line = lines[time_line_index]
        parts = time_line.split("-->")
        if len(parts) != 2:
            continue

        start = parts[0].strip()
        end = parts[1].strip()
        text_lines = lines[time_line_index + 1:]
        cue_text = "\n".join(text_lines).strip()
        if cue_text:
            cues.append({"start": start, "end": end, "text": cue_text})
    return cues

# =====================================================
# 時刻変換系（Decimalで精度維持し、フレームは0〜fps-1に正規化）
# =====================================================

def vtt_time_to_tc_string(vtt_time: str, fps: int = FPS) -> str:
    """
    "00:00:04.288" → "00:00:04:09" のような FCP向けタイムコード文字列に変換。
    フレーム計算は Decimal で行い、切り上げでfpsに達したら秒を+1してフレーム0に繰り上げる。
    """
    h, m, s = vtt_time.split(":")

    # 00:00:00.000 の場合は 00:00:00.100 に変換
    if h == "00" and m == "00" and s == "00.000":
        s = "00.100"
        print("00:00:00.000 は 00:00:00.100 に変換されました")

    sec_dec = Decimal(s)
    total_sec_dec = Decimal(int(h)) * 3600 + Decimal(int(m)) * 60 + sec_dec

    # 整数秒は切り捨て、残りをフレーム化
    sec_int = int(total_sec_dec)  # floor
    frac = total_sec_dec - Decimal(sec_int)

    frame_dec = (frac * fps).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    frame = int(frame_dec)

    if frame >= fps:
        frame = 0
        sec_int += 1

    h_out = sec_int // 3600
    m_out = (sec_int % 3600) // 60
    s_out = sec_int % 60

    return f"{h_out:02d}:{m_out:02d}:{s_out:02d}:{frame:02d}"

def vtt_time_to_seconds(vtt_time: str) -> float:
    """"00:00:04.288" → 秒数(float)に変換。Decimalで精度維持。"""
    h, m, s = vtt_time.split(":")
    sec_dec = Decimal(s)
    total_sec_dec = Decimal(int(h)) * 3600 + Decimal(int(m)) * 60 + sec_dec
    return float(total_sec_dec)

def add_offset_to_vtt_time(vtt_time: str, offset_sec: float = 0.1) -> str:
    """
    VTT形式の時刻に offset_sec 秒を加算した新しい VTT 時刻文字列を返す。
    ミリ秒まで四捨五入し、桁あふれがあれば繰り上げる。
    """
    h, m, s = vtt_time.split(":")
    sec_dec = Decimal(s)
    base_ms = (Decimal(int(h)) * 3600 + Decimal(int(m)) * 60 + sec_dec) * Decimal(1000)
    offset_ms = (Decimal(str(offset_sec)) * Decimal(1000)).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    total_ms = (base_ms + offset_ms).quantize(Decimal("1"), rounding=ROUND_HALF_UP)

    total_ms_int = int(total_ms)
    h_out = total_ms_int // 3_600_000
    rem = total_ms_int % 3_600_000
    m_out = rem // 60_000
    rem = rem % 60_000
    s_out = rem // 1_000
    ms_out = rem % 1_000

    return f"{h_out:02d}:{m_out:02d}:{s_out:02d}.{ms_out:03d}"

def move_playhead_to_time(vtt_time: str):
    """
    指定 VTT 時刻に再生ヘッドを移動する。
    ctrl+p → Cmd+V → Enter
    """
    tc_str = vtt_time_to_tc_string(vtt_time)
    print(f"[INFO] Move playhead to {vtt_time} (TC: {tc_str})")

    pyautogui.keyDown("ctrl")
    pyautogui.press("p")
    pyautogui.keyUp("ctrl")
    time.sleep(SLEEP_SHORT)

    pyperclip.copy(tc_str)
    pyautogui.hotkey("command", "v")
    time.sleep(SLEEP_SHORT)

    pyautogui.press("enter")
    time.sleep(SLEEP_SHORT)

# =====================================================
# カットポイント抽出
# =====================================================

def collect_cut_points(cues):
    """開始・終了時刻を集め、秒でソートし重複除去したリストを返す。"""
    points = []
    for cue in cues:
        points.append((vtt_time_to_seconds(cue["start"]), cue["start"]))
        points.append((vtt_time_to_seconds(cue["end"]), cue["end"]))

    points_sorted = sorted(points, key=lambda x: x[0])
    cut_entries = []
    seen = set()

    for _, t in points_sorted:
        adjusted = t
        adjusted_sec = vtt_time_to_seconds(adjusted)

        # 直前のカットと0.1秒以下なら押し出し、重複時刻も避ける
        if cut_entries:
            last_sec = cut_entries[-1][0]
            while adjusted_sec - last_sec <= MIN_GAP_SEC or adjusted in seen:
                adjusted = add_offset_to_vtt_time(adjusted, offset_sec=MIN_GAP_SEC)
                adjusted_sec = vtt_time_to_seconds(adjusted)
        else:
            while adjusted in seen:
                adjusted = add_offset_to_vtt_time(adjusted, offset_sec=MIN_GAP_SEC)
                adjusted_sec = vtt_time_to_seconds(adjusted)

        seen.add(adjusted)
        cut_entries.append((adjusted_sec, adjusted))

    # 微調整により秒数が増えた場合も昇順を維持する
    cut_entries.sort(key=lambda x: x[0])
    return [t for _, t in cut_entries]

# =====================================================
# 編集操作系
# =====================================================

def blade_at_playhead():
    """再生ヘッド位置でクリップを分割 (command + B)"""
    pyautogui.keyDown("command")
    pyautogui.press("b")
    pyautogui.keyUp("command")
    time.sleep(SLEEP_SHORT)

def go_to_next_clip():
    """次のテキストクリップへ移動（command + right）"""
    pyautogui.keyDown("command")
    pyautogui.press("right")
    pyautogui.keyUp("command")
    time.sleep(SLEEP_CLIP_MOVE)
    print("command+right")

def focus_text_field_first_time():
    """インスペクタ内のテキストフィールドをクリックしてフォーカスを当てる"""
    pyautogui.click(INPUT_X, INPUT_Y)
    time.sleep(SLEEP_SHORT)

def paste_text(text: str):
    """テキストフィールドに text をペースト（Cmd+V）"""
    pyperclip.copy(text)
    pyautogui.hotkey("command", "v")
    time.sleep(SLEEP_SHORT)
    print("command+v")
    print("text", text)

# =====================================================
# メイン処理
# =====================================================

def main():
    # 1. VTT を解析
    cues = parse_vtt_from_file(VTT_FILE)
    if not cues:
        print(f"⚠ VTTからキューが1件も読み取れませんでした: {VTT_FILE}")
        return

    cut_points = collect_cut_points(cues)
    if not cut_points:
        print("⚠ カットポイントが抽出できませんでした。VTTを確認してください。")
        return

    print(f"📄 VTT ファイル: {VTT_FILE}")
    print("=== 解析結果 ===")
    for c in cues:
        print(c)
    print("================")

    print(f"⏱ {SLEEP_COUNTDOWN}秒後に開始します。Final Cut Pro をアクティブにし、")
    print("テキストクリップのみタイムラインに置く。テキストクリップの先頭の位置を登録しておく。\nShift+Zでテキストクリップ全体が見えるようにしておく。\nタイムライン上の長いテキストクリップ（入力テキストは空にしておく）を1本選択した状態にしてください。")
    time.sleep(SLEEP_COUNTDOWN)

    # 2. カット
    print("✂ テキストクリップをカットします（貼り付けなし）...")
    for t in cut_points:
        print(f"  - Cut at {t}")
        move_playhead_to_time(t)
        blade_at_playhead()

    print("✅ カット完了。ここから貼り付けを行います。")

    # 2つ目のテキストクリップへ移動
    go_to_next_clip()

    # テキストフィールドをマウスクリックで選択（初回のみ）
    focus_text_field_first_time()
    time.sleep(SLEEP_CLIP_MOVE)

    # 1つ目のセリフを貼り付け
    paste_text(cues[0]["text"])
    time.sleep(SLEEP_CLIP_MOVE)

    # 2つ目以降のセリフを貼り付け
    for cue in cues[1:]:
        # セリフクリップの間に無音クリップがあるため、2つ進む
        go_to_next_clip()
        go_to_next_clip()
        pyautogui.press("tab")
        time.sleep(SLEEP_CLIP_MOVE)
        paste_text(cue["text"])
        time.sleep(SLEEP_CLIP_MOVE)

    print("✅ 貼り付けまで完了しました。")
    print("🎉 すべて完了しました。")


if __name__ == "__main__":
    main()
