import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import sys
import json

# 出力をUTF-8に設定
if sys.stdout.encoding != 'utf-8':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def get_mercari_profile(user_id):
    profile_url = f"https://jp.mercari.com/user/profile/{user_id}"

    # オプション設定
    options = uc.ChromeOptions()
    # options.add_argument('--headless') # 一旦コメントアウト
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')

    # ユーザーエージェントを人間っぽく装飾（大事！）
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36')

    print(f"[*] メルカリに潜入中: {user_id}")

    try:
        driver = uc.Chrome(options=options)
        driver.get(profile_url)

        # 読み込み待ち
        wait = WebDriverWait(driver, 20)

        # データの抽出 (2026年時点のセレクタを想定)
        result = {}

        # ユーザー名
        result['user_name'] = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'h1[data-testid="profile-name"]'))).text

        # 評価数
        try:
            result['ratings'] = driver.find_element(By.CSS_SELECTOR, 'span[data-testid="ratings-count"]').text
        except: result['ratings'] = "N/A"

        # 出品数
        try:
            result['listings'] = driver.find_element(By.CSS_SELECTOR, 'span[data-testid="listings-count"]').text
        except: result['listings'] = "N/A"

        # 自己紹介文 (最初の部分のみ)
        try:
            result['intro'] = driver.find_element(By.CSS_SELECTOR, 'div[data-testid="profile-description"]').text[:50].replace('\n', ' ') + "..."
        except: result['intro'] = "N/A"

        return result

    except Exception as e:
        print(f"[!] エラー発生: {e}")
        return None
    finally:
        try: driver.quit()
        except: pass

if __name__ == "__main__":
    target_user = "491901537" # ストライプさん
    profile = get_mercari_profile(target_user)

    if profile:
        print("\n--- 調査結果 (ASSC v1.0連携) ---")
        print(f"【ユーザー名】: {profile['user_name']}")
        print(f"【評価数】    : {profile['ratings']}")
        print(f"【出品数】    : {profile['listings']}")
        print(f"【自己紹介】  : {profile['intro']}")
        print("--------------------------------")
    else:
        print("[X] 情報の取得に失敗しました。")
