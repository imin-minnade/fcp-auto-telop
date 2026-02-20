"""WAV ファイルのファイル名を「セリフ_番号.wav」→「番号_セリフ.wav」にリネームするスクリプト。

auto_aques_talk_player.py で生成した音声ファイルは「セリフ_番号.wav」という命名になっている。
これを「番号_セリフ.wav」にリネームすることで、
Final Cut Pro に読み込んだ際に正しい順番で並ぶようになる。

使い方:
1. wav_output/ フォルダに対象の WAV ファイルを配置
2. `python scripts/swap_title_number.py` を実行
"""

import os

# === 設定 ===
folder_path = "./wav_output"

# === メイン処理 ===
for filename in os.listdir(folder_path):
    # WAVファイルのみ処理
    if filename.endswith(".wav"):
        name_part = filename[:-4]

        # アンダースコアで分割
        if "_" in name_part:
            title, number = name_part.rsplit("_", 1)

            # 数字部分が整数として成立する場合のみリネーム
            if number.isdigit():
                new_name = f"{number}_{title}.wav"

                # パスを組み立て
                old_path = os.path.join(folder_path, filename)
                new_path = os.path.join(folder_path, new_name)

                # リネーム実行
                os.rename(old_path, new_path)
                print(f"✅ {filename} → {new_name}")
            else:
                print(f"⚠️ スキップ: 数字が見つかりません → {filename}")
        else:
            print(f"⚠️ スキップ: '_' が含まれていません → {filename}")

print("完了しました。")
