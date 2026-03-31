import os
import sys
import json

def load_json_safely(path):
    """BOM付きUTF-8、通常のUTF-8、Shift-JISを順番に試す賢い読み込みクラス"""
    encodings = ['utf-8-sig', 'utf-8', 'cp932']
    for enc in encodings:
        try:
            with open(path, "r", encoding=enc) as f:
                return json.load(f)
        except (UnicodeDecodeError, json.JSONDecodeError):
            continue
    return None

def create_mercari_folder(root_path, folder_name):
    mercari_path = os.path.join(root_path, "mercari")
    new_folder_path = os.path.join(mercari_path, folder_name)
    original_path = os.path.join(new_folder_path, "Original")

    # config.jsonのパス（スクリプトと同じ場所）
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")     

    try:
        # 1. フォルダ作成
        os.makedirs(original_path, exist_ok=True)

        # 2. config.jsonの読み込み
        desc_text = "(説明文データが見つかりませんでしたッ！)"
        config_list = load_json_safely(config_path)

        if config_list:
            # JSONから説明文(DESC)を探す
            for item in config_list:
                if isinstance(item, dict) and item.get("name") == "DESC":
                    desc_text = item.get("text", "")
                    break
        else:
            print("【警告】config.jsonが読み込めないか、または形式が不正です。")

        # 3. 説明文テキストの保存
        txt_name = f"{folder_name[:4]}-説明文.txt"
        txt_path = os.path.join(new_folder_path, txt_name)
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(f"[件名]\n{folder_name}\n\n\n[本文]\n{desc_text}")

        print(f"========================================")
        print(f"フォルダと説明文を完全に用意したぜッ！（はーと）")
        print(f"場所: {new_folder_path}")
        print(f"========================================")

        # 4. フォルダを自動で開く
        os.startfile(new_folder_path)

    except Exception as e:
        print(f"【エラー】: {e}")

if __name__ == "__main__":
    if len(sys.argv) >= 3:
        create_mercari_folder(sys.argv[1], sys.argv[2])
    else:
        print("引数が足りないぜッ！ [ルートパス] [フォルダ名] を指定してッマセ！")
