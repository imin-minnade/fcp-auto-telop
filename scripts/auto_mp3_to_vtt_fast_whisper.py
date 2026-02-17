"""faster-whisper で音声/動画ファイルを文字起こしし、VTT に保存するスクリプト。

使い方:
1. audio_input/ に音声ファイル（mp3/mp4/m4a/wav）を配置
2. AUDIO_FILENAME を対象ファイル名に変更
3. `pip install -r requirements.txt`（初回はモデル自動ダウンロード）
4. `python scripts/auto_mp3_to_vtt_fast_whisper.py` を実行
5. vtt_output/ に .vtt ファイルが出力されます
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable, List, Tuple

from faster_whisper import WhisperModel

# ===================== 設定 =====================
# 音声ファイルの入力ディレクトリ
AUDIO_DIR = Path("audio_input")

# VTT ファイルの出力ディレクトリ
VTT_DIR = Path("vtt_output")

# 対象ファイル名（拡張子なしの場合は自動探索）
AUDIO_FILENAME = "sample.mp3"

# 探索対象の拡張子（優先順）
ALLOWED_EXTS = [".mp3", ".mp4", ".m4a", ".mov", ".wav"]

# Whisper モデル名（精度優先: "large-v2" / 速度優先: "small"）
MODEL_NAME = "large-v2"

# デバイス: "auto" / "cpu" / "cuda" / "metal"（Mac の場合は auto で OK）
DEVICE = "auto"

# 計算精度: "auto" / "float16" / "int8_float16" など
COMPUTE_TYPE = "auto"

# 句読点として認識する文字
SENTENCE_PUNCT = "。.!?！？"

# 文を分ける無音ギャップのしきい値（秒）
MAX_GAP_SECONDS = 0.2

# 句読点がなくても強制分割する長さ（秒）。None で無効
MAX_SENTENCE_SECONDS: float | None = None

# 句読点がなくても強制分割する単語数。None で無効
MAX_SENTENCE_WORDS: int | None = None
# ===================== 設定ここまで =====================


def format_timestamp(seconds: float) -> str:
    """秒数を VTT タイムスタンプ形式に変換する。"""
    millis = int(round(seconds * 1000))
    hours, remainder = divmod(millis, 3600 * 1000)
    minutes, remainder = divmod(remainder, 60 * 1000)
    secs, ms = divmod(remainder, 1000)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}.{ms:03d}"


def resolve_audio_path(filename: str, search_dir: Path, allowed_exts: List[str]) -> Path:
    """音声ファイルのパスを解決する。拡張子なしの場合は自動探索。"""
    base_path = search_dir / filename

    if base_path.suffix:
        if base_path.exists():
            return base_path
        stem = base_path.stem
    else:
        stem = filename

    for ext in allowed_exts:
        candidate = search_dir / f"{stem}{ext}"
        if candidate.exists():
            return candidate

    exts_str = ", ".join(allowed_exts)
    raise FileNotFoundError(f"音声/動画ファイルが見つかりません: {stem} ({exts_str})")


def transcribe_to_vtt(audio_path: Path) -> str:
    """音声ファイルを文字起こしして VTT 形式のテキストを返す。"""
    if not audio_path.exists():
        raise FileNotFoundError(f"音声ファイルが見つかりません: {audio_path}")

    model = WhisperModel(MODEL_NAME, device=DEVICE, compute_type=COMPUTE_TYPE)
    segments, _info = model.transcribe(
        str(audio_path),
        beam_size=5,
        word_timestamps=True,
        vad_filter=False,
    )

    sentence_segments = split_into_sentences(segments)
    lines: list[str] = ["WEBVTT", ""]
    for start_sec, end_sec, text in sentence_segments:
        lines.append(f"{format_timestamp(start_sec)} --> {format_timestamp(end_sec)}")
        lines.append(text)
        lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def split_into_sentences(segments: Iterable) -> List[Tuple[float, float, str]]:
    """Whisper のセグメントを句読点・無音ギャップで文単位に分割する。"""
    sentences: List[Tuple[float, float, str]] = []

    for seg in segments:
        words = seg.words or []
        if not words:
            sentences.append((seg.start, seg.end, seg.text.strip()))
            continue

        buffer: list[str] = []
        start_time = words[0].start
        last_end = words[0].end

        for w in words:
            word_text = w.word
            start_w, end_w = w.start, w.end

            # 無音ギャップで分割
            if buffer and (start_w - last_end) > MAX_GAP_SECONDS:
                sentences.append((start_time, last_end, "".join(buffer).strip()))
                buffer = []
                start_time = start_w

            buffer.append(word_text)
            last_end = end_w

            punct_hit = word_text.strip() and word_text.strip()[-1] in SENTENCE_PUNCT
            over_time = (
                MAX_SENTENCE_SECONDS is not None
                and start_time is not None
                and (last_end - start_time) >= MAX_SENTENCE_SECONDS
            )
            over_words = MAX_SENTENCE_WORDS is not None and len(buffer) >= MAX_SENTENCE_WORDS

            if punct_hit or over_time or over_words:
                sentences.append((start_time, last_end, "".join(buffer).strip()))
                buffer = []
                start_time = None

        if buffer:
            sentences.append((start_time or words[0].start, last_end, "".join(buffer).strip()))

    return sentences


def save_vtt(vtt_text: str, output_path: Path) -> None:
    """VTT テキストをファイルに保存する。"""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(vtt_text, encoding="utf-8")


def main() -> None:
    audio_path = resolve_audio_path(AUDIO_FILENAME, AUDIO_DIR, ALLOWED_EXTS)
    output_path = VTT_DIR / audio_path.with_suffix(".vtt").name

    print(f"🎙 文字起こし開始: {audio_path}")
    print(f"📂 出力先: {output_path}")
    vtt_text = transcribe_to_vtt(audio_path)
    save_vtt(vtt_text, output_path)
    print(f"✅ 完了: {output_path}")


if __name__ == "__main__":
    main()
