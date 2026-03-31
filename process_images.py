import sys
import os
import json
from PIL import Image, ImageDraw, ImageFont

# フォント設定の定数化
DEFAULT_FONT_NAME = "yugothm.ttc"
FONT_SEARCH_PATHS = [
    "C:/Windows/Fonts/" + DEFAULT_FONT_NAME,
    "/System/Library/Fonts/Hiragino Sans GB.ttc", # Mac用予備
    "C:/Windows/Fonts/msgothic.ttc"               # 予備
]

def get_font(size):
    """日本語対応フォントを安全に取得する"""
    for path in FONT_SEARCH_PATHS:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except:
                continue
    return ImageFont.load_default()

def draw_text_with_offset(draw, cfg):
    """設定に基づき画像に文字を描画する共通処理"""
    label = cfg.get("name", "")
    if not label or label == "SKIP":
        return

    # 改行位置の制御
    display_text = label
    offset = int(cfg.get("offset", 0))
    if 0 < offset < len(label):
        display_text = label[:offset] + "\n" + label[offset:]

    font = get_font(cfg.get("size", 20))
    color = tuple(cfg.get("color", [255, 255, 255]))
    pos = (cfg.get("x", 0), cfg.get("y", 0))

    draw.multiline_text(
        pos, display_text, fill=color, font=font,
        align="center", anchor="mm", spacing=cfg.get("size", 20) * 0.2
    )

def process_logic(root_path, folder_name, config_list, is_preview=False):
    root_path = os.path.normpath(root_path)
    target_dir = os.path.join(root_path, "mercari", folder_name, "Original")
    output_dir = os.path.join(root_path, "mercari", folder_name)

    if not os.path.exists(target_dir):
        print(f"致命的エラー: フォルダが見つからない -> {target_dir}")
        return

    valid_ext = ('.jpg', '.jpeg', '.png')
    files = sorted([f for f in os.listdir(target_dir) if f.lower().endswith(valid_ext)])      

    if not files:
        print("処理対象の画像ファイルがありません。")
        return

    # --- 1枚目：集合写真処理 ---
    with Image.open(os.path.join(target_dir, files[0])).convert("RGB") as img:
        draw = ImageDraw.Draw(img)
        # JSONの最初の7要素を集合写真用としてループ
        for i in range(min(8, len(config_list))):
            draw_text_with_offset(draw, config_list[i])

        save_name = "PREVIEW_CHECK.jpg" if is_preview else files[0]
        save_path = os.path.join(output_dir, save_name)
        img.save(save_path, quality=95)
        print(f"Success: 1枚目保存完了 -> {save_name}")

        if is_preview:
            os.startfile(save_path)
            return

    # --- 2枚目以降：本番一括モード ---
    # 2枚目（インデックス1）は無加工コピー
    if len(files) >= 2:
        with Image.open(os.path.join(target_dir, files[1])) as img:
            img.save(os.path.join(output_dir, files[1]))

    # 3枚目以降に個別の名前を入れる
    for i in range(2, min(len(files), 9)):
        cfg_idx = i + 6 # 3枚目(i=2)なら config[7] を参照する仕様を維持
        if cfg_idx >= len(config_list): break

        with Image.open(os.path.join(target_dir, files[i])).convert("RGB") as img:
            draw = ImageDraw.Draw(img)
            draw_text_with_offset(draw, config_list[cfg_idx])

            # リネーム処理
            base, ext = os.path.splitext(files[i])
            label = config_list[cfg_idx].get("name", "")
            out_name = f"{base}_{label}{ext}" if label and label != "SKIP" else files[i]      

            img.save(os.path.join(output_dir, out_name), quality=95)
            print(f"Success: {i+1}枚目保存完了 -> {out_name}")

if __name__ == "__main__":
    # 1. 引数のチェック
    if len(sys.argv) < 3:
        print("【エラー】引数が足りないぜッ！ [root_path] [folder_name] を指定してくれ。")    
        sys.exit(1)

    r_path = sys.argv[1]
    f_name = sys.argv[2]
    # 第3引数に "Full" があれば本番モード
    mode_full = (len(sys.argv) > 3 and sys.argv[3] == "Full")

    # 2. config.json のパス設定
    # スクリプトと同じ場所にある config.json を探す
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, "config.json")

    # 3. JSONの読み込み（より堅牢なバージョン）
    conf = None
    if os.path.exists(config_path):
        # 読み込みを試行するエンコード順リスト
        # utf-8-sig はBOM付きUTF-8を自動判別するので最優先
        encodings = ['utf-8-sig', 'utf-8', 'cp932']

        for enc in encodings:
            try:
                with open(config_path, "r", encoding=enc) as f:
                    conf = json.load(f)
                # 成功したらループを抜ける
                print(f"DEBUG: {enc} で設定ファイルを読み込んだぞッ！")
                break
            except Exception:
                continue

        if conf is None:
            print(f"【エラー】どの文字コードでもJSONを解釈できなかったぜッ！...")
    else:
        print(f"【エラー】config.json が見つからないぜッ！：{config_path}")

    # 4. 実行：conf が無事に読み込めていれば実行するッ！
    if conf:
        process_logic(r_path, f_name, conf, is_preview=not mode_full)
    else:
        print("【中止】設定データがないため、処理を中断したぜッ！")
