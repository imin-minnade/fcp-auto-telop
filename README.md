# FCP Auto Telop 🎬

**Final Cut Pro のテロップ作業を自動化する Python ツールキット**

動画1本あたりのテロップ作業を **約75%短縮** します。
合成音声（VOICEVOX / AquesTalk）と生声収録の両方に対応。

![Python](https://img.shields.io/badge/Python-3.9%2B-blue)
![Platform](https://img.shields.io/badge/Platform-macOS-lightgrey)
![License](https://img.shields.io/badge/License-MIT-green)

---

## どんなツール？

YouTube動画のテロップ（テキストクリップ）挿入を自動化するスクリプト集です。

Final Cut Pro のタイムライン上で直接テキストクリップを分割・貼り付けするため、作業過程を視覚的に確認しながら安心して自動化を進められます。

### 対応ワークフロー

| ワークフロー | 説明 |
|---|---|
| **合成音声ルート** | シナリオ（CSV）→ AquesTalk で音声生成 → FCP でテロップ自動挿入 |
| **生声ルート** | 録音 → Whisper で文字起こし（VTT）→ FCP でテロップ自動挿入 |

### なぜ VREW や AI キャプションではなく、このツール？

- **Final Cut Pro 上で直接編集可能**: テキストクリップとして挿入されるため、フォント・サイズ・色・アニメーションを自由にカスタマイズできます
- **AI キャプションの制限を回避**: FCP の AI キャプションは日本語未対応（2025年時点）であり、対応後もスタイル変更ができません。本ツールは SRT/VTT からテキストクリップへの変換に対応しています
- **環境に依存しない**: キャリブレーションスクリプトで座標を自動取得するため、どの Mac 環境でも動作します

---

## ディレクトリ構成

```
fcp-auto-telop/
├── scripts/                    # 自動化スクリプト
│   ├── auto_fcp_telop_split_paste.py
│   ├── auto_fcp_vtt_to_telop.py
│   ├── auto_audio_to_vtt.py
│   ├── auto_aques_talk_player.py
│   ├── swap_title_number.py
│   ├── vtt_timestamp_checker.py
│   └── get_mouse_positions.py
├── csv_input/                  # シナリオ CSV を配置
│   └── sample.csv
├── txt_input/                  # セリフ TXT を配置
│   └── sample.txt
├── audio_input/                # 音声ファイルを配置（生声ルート）
├── vtt_input/                  # VTT ファイルを配置
│   └── sample.vtt
├── vtt_output/                 # Whisper の文字起こし結果が出力される
├── wav_output/                 # AquesTalk の音声ファイルを配置
├── requirements.txt
└── README.md
```

## スクリプト一覧

| スクリプト | 機能 | 入力 | 出力 |
|---|---|---|---|
| `auto_fcp_telop_split_paste.py` | FCP 上でテキストクリップの分割とセリフの貼り付け | `txt_input/*.txt` | FCP タイムライン |
| `auto_fcp_vtt_to_telop.py` | VTT の時刻でクリップを分割し、セリフを貼り付け | `vtt_input/*.vtt` | FCP タイムライン |
| `auto_audio_to_vtt.py` | faster-whisper で音声ファイルを文字起こし | `audio_input/*` | `vtt_output/*.vtt` |
| `auto_aques_talk_player.py` | AquesTalk Player への読み上げテキスト自動入力 | `csv_input/*.csv` | `wav_output/*.wav` |
| `swap_title_number.py` | WAV ファイル名を「セリフ_番号」→「番号_セリフ」にリネーム | `wav_output/*.wav` | `wav_output/*.wav` |
| `vtt_timestamp_checker.py` | VTT のタイムスタンプ重なりをチェック | `vtt_input/*.vtt` | 標準出力 |
| `get_mouse_positions.py` | クリックした画面座標を表示 | マウス操作 | 座標値の表示 |

---

## クイックスタート

### 1. 環境構築

```bash
git clone https://github.com/YOUR_USERNAME/fcp-auto-telop.git
cd fcp-auto-telop
pip install -r requirements.txt
```

### 2. 座標キャリブレーション

お使いの Mac の画面解像度に合わせて、FCP のテキスト入力欄の座標を取得します。

```bash
python scripts/get_mouse_positions.py
```

取得した座標を各スクリプトの `INPUT_X, INPUT_Y` に設定してください。

### 3A. 合成音声ルート

1. `csv_input/` にシナリオ CSV を配置（`sample.csv` を参照）
2. AquesTalk Player で音声を自動生成：

```bash
python scripts/auto_aques_talk_player.py
```

3. `wav_output/` の WAV ファイルをリネーム（FCP で正しい順番に並ぶよう番号を先頭に）：

```bash
python scripts/swap_title_number.py
```

4. WAV ファイルを FCP に読み込み、テキストクリップを選択した状態にする
5. テロップを自動挿入：

```bash
python scripts/auto_fcp_telop_split_paste.py
```

### 3B. 生声ルート

1. 音声ファイルを `audio_input/` に配置
2. Whisper で文字起こし：

```bash
python scripts/auto_audio_to_vtt.py
```

3. `vtt_output/` に生成された VTT を `vtt_input/` にコピーして確認・修正
4. タイムスタンプの重なりをチェック：

```bash
python scripts/vtt_timestamp_checker.py
```

5. FCP でテロップを自動挿入：

```bash
python scripts/auto_fcp_vtt_to_telop.py
```

---

## CSV フォーマット（合成音声ルート）

`auto_aques_talk_player.py` で使用します。Numbers や Excel で編集できます。

| 列名 | 説明 | 例 |
|---|---|---|
| `実行` | 1 = 実行、0 = スキップ | `1` |
| `キャラクター` | キャラクター名（フィルタリング用） | `魔理沙` |
| `セリフ` | 読み上げるテキスト | `こんにちは` |

```csv
実行,キャラクター,セリフ
1,魔理沙,皆さんこんにちは。
1,魔理沙,今回は動画制作の効率化について紹介します。
0,魔理沙,このセリフはスキップされます。
```

---

## 必要環境

- macOS（Final Cut Pro が動作する環境）
- Python 3.9 以上
- Final Cut Pro（テキストクリップのテンプレートを事前に配置）

---

## 詳しい使い方・ワークフロー解説

各スクリプトの詳しい設定方法、テンプレートの作り方、312本の動画制作で培った効率化テクニックを解説した記事を公開しています。

👉 **[Final Cut Pro × Python 動画制作自動化バイブル（note）](https://note.com/YOUR_NOTE_URL)**

---

## ライセンス

MIT License - 個人利用・商用利用ともに自由にお使いいただけます。

---

## 作者

YouTube チャンネルで実際にこのツールを使って動画を制作しています。

- YouTube: [チャンネルリンク](https://www.youtube.com/channel/UCPKMH4pEKMhZt7cNGfjoSRQ)
- X (Twitter): [@YOUR_HANDLE](https://x.com/YOUR_HANDLE)
- note: [YOUR_NOTE_URL](https://note.com/YOUR_NOTE_URL)
