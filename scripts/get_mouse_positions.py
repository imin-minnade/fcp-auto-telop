"""マウスクリックした座標をコンソールに表示するスクリプト。

使い方:
1. このスクリプトを実行
2. 座標を取得したい位置をクリック
3. コンソールに座標が表示される
4. Esc キーで終了
"""

from pynput import keyboard, mouse


def on_click(x, y, _button, pressed):
    if pressed:
        print(f"{int(x)}, {int(y)}")


def on_press(key):
    if key == keyboard.Key.esc:
        mouse_listener.stop()
        return False


mouse_listener = mouse.Listener(on_click=on_click)


print("クリックした座標を表示します。Esc キーで終了。")
print()

mouse_listener.start()

with keyboard.Listener(on_press=on_press) as kb_listener:
    kb_listener.join()
