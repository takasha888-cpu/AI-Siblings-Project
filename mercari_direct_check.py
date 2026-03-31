import requests
import re

user_id = "491901537"
url = f"https://jp.mercari.com/user/profile/{user_id}"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept-Language": "ja-JP,ja;q=0.9"
}

print(f"[*] 直接リクエストを試行中...")
try:
    response = requests.get(url, headers=headers, timeout=15)
    html = response.text

    # ユーザー名を探す
    name_match = re.search(r'\"name\":\"([^\"]+)\"', html)
    # 評価数
    rating_match = re.search(r'\"num_ratings\":(\d+)', html)
    # 出品数
    listing_match = re.search(r'\"num_sell_items\":(\d+)', html)

    if name_match:
        print(f"\n--- 取得成功！ ---")
        print(f"ユーザー名: {name_match.group(1)}")
        if rating_match: print(f"評価数: {rating_match.group(1)}")
        if listing_match: print(f"出品数: {listing_match.group(1)}")
        print("------------------")
    else:
        print("[!] ページは見えましたが、詳細データが隠されています。")
        # フッターしか取れなかった場合など
except Exception as e:
    print(f"Error: {e}")
