"""
VTT ファイルのタイムスタンプ重なりをチェックするスクリプト。
前の終了時刻より次の開始時刻が早い（gap < 0）箇所を報告します。
問題がなければ「タイムスタンプに異常はありません」と表示します。

使い方:
1. VTT_FILE を対象ファイル名に変更
2. `python scripts/vtt_timestamp_checker.py` を実行
"""

import os
import sys
from typing import List, Tuple

# ===================== 設定 =====================
# チェック対象の VTT ファイル（vtt_input/ ディレクトリに配置）
VTT_FILE = "vtt_input/sample.vtt"
# ===================== 設定ここまで =====================


def load_vtt(vtt_path: str) -> str:
    """VTT ファイルを読み込んでテキストを返す。"""
    if not os.path.exists(vtt_path):
        print(f"❌ VTT ファイルが見つかりません: {vtt_path}")
        sys.exit(1)
    with open(vtt_path, encoding="utf-8") as f:
        return f.read()


def to_seconds(timestamp: str) -> float:
    """hh:mm:ss.mmm を秒(float)に変換"""
    hms, millis = timestamp.split(".")
    hours, minutes, seconds = map(int, hms.split(":"))
    return hours * 3600 + minutes * 60 + seconds + float(f"0.{millis}")


def parse_segments(text: str) -> List[Tuple[float, float, int, str]]:
    """タイムレンジ行だけを抽出して (start, end, line_no, line_text) に変換"""
    segments: List[Tuple[float, float, int, str]] = []
    for idx, line in enumerate(text.splitlines(), start=1):
        if "-->" not in line:
            continue
        try:
            start_str, end_str = line.split("-->")
            start = to_seconds(start_str.strip())
            end = to_seconds(end_str.strip())
        except Exception as exc:  # noqa: BLE001
            print(f"警告: {idx}行目を解析できませんでした: {line} ({exc})")
            continue
        segments.append((start, end, idx, line.strip()))
    return segments


def check_intervals(segments: List[Tuple[float, float, int, str]]) -> None:
    """前区間の終了より次区間の開始が早い（重なり）箇所を報告"""
    issues = []
    for i in range(1, len(segments)):
        prev_start, prev_end, prev_line_no, prev_text = segments[i - 1]
        curr_start, curr_end, curr_line_no, curr_text = segments[i]
        gap = curr_start - prev_end
        if gap < 0:
            issues.append(
                {
                    "gap": gap,
                    "prev_line_no": prev_line_no,
                    "curr_line_no": curr_line_no,
                    "prev_text": prev_text,
                    "curr_text": curr_text,
                    "prev_range": (prev_start, prev_end),
                    "curr_range": (curr_start, curr_end),
                }
            )

    if not issues:
        print("タイムスタンプに異常はありません")
        return

    for issue in issues:
        print("----")
        print(
            f"重なりあり (差: {issue['gap']:.3f} 秒)"
            f" | 前: {issue['prev_range'][0]:.3f}-{issue['prev_range'][1]:.3f}"
            f" -> 次: {issue['curr_range'][0]:.3f}-{issue['curr_range'][1]:.3f}"
        )
        print(f"前の行(L{issue['prev_line_no']}): {issue['prev_text']}")
        print(f"次の行(L{issue['curr_line_no']}): {issue['curr_text']}")


def main() -> None:
    print(f"📄 VTT ファイル: {VTT_FILE}")
    vtt_text = load_vtt(VTT_FILE)
    segments = parse_segments(vtt_text)
    if not segments:
        print("タイムスタンプ行が見つかりませんでした")
        return
    check_intervals(segments)


if __name__ == "__main__":
    main()
