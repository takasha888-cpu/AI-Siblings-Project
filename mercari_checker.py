import urllib.request
import json
import re
import sys

def get_mercari_user(user_id):
    url = f"https://jp.mercari.com/user/profile/{user_id}"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
        'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8',
    }

    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req) as response:
            html = response.read().decode('utf-8')

            # __NEXT_DATA__ を探す
            match = re.search(r'<script id="__NEXT_DATA__" type="application/json">(.*?)</script>', html)
            if not match:
                return "Error: Could not find __NEXT_DATA__"

            data = json.loads(match.group(1))
            apollo_state = data.get('props', {}).get('pageProps', {}).get('apolloState', {})  

            # User: から始まるキーを探す
            user_key = next((k for k in apollo_state.keys() if k.startswith(f"User:{user_id}")), None)
            if not user_key:
                # ユーザーIDが分からない場合は、最初に見つかった User: キーを使う
                user_key = next((k for k in apollo_state.keys() if k.startswith("User:")), None)

            if user_key:
                user_data = apollo_state[user_key]
                name = user_data.get('name', 'Unknown')
                ratings = user_data.get('ratingsCount', 0)
                items = user_data.get('itemsCount', 0)
                # フォロワー数
                followers = user_data.get('followersCount', 0)
                # 出品者レベル（ランク）
                seller_rank = user_data.get('sellerRank', 0)

                return f"SUCCESS|{name}|{seller_rank}|{ratings}|{items}|{followers}"
            else:
                return "Error: User key not found in apollo state"

    except Exception as e:
        return f"Error: {str(e)}"

if __name__ == "__main__":
    # たにくるさんのID: 897999199
    result = get_mercari_user("897999199")
    print(result)
