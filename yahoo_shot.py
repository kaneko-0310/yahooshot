# -*- coding: utf-8 -*-
"""
Yahoo! 検索サジェスト自動撮影ツール（修正版）
Google Spreadsheetsからキーワード取得 → Yahoo検索でサジェスト表示時に即座撮影
"""
import os
import time
import csv
import io
import requests
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

# ========= 設定 =========
# Google Sheets URL (CSVエクスポート用)
SPREADSHEET_ID = "1i2SKztLstWUeD0y9mNsNtdYrwdtZ7zCzNDA7vhojA8s"
SHEET_GID = "1060586181"  # シートのGID
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/export?format=csv&gid={SHEET_GID}"

OUTPUT_DIR = r"C:\Users\kanek\Pictures\朝"
YAHOO_URL = "https://www.yahoo.co.jp/"

# iPhone 16の画面サイズ
VIEWPORT_W, VIEWPORT_H = 393, 852
USE_INCOGNITO = True

# タイミング調整
TYPE_INTERVAL = 0.3      # 文字入力間隔（少し速く）
PAUSE_AFTER_TYPE = 1.0   # 入力完了後の待機
SCREENSHOT_DELAY = 0.5   # サジェスト検出後の撮影遅延
# =======================

def get_keywords_from_sheets():
    """Google SheetsからE列のキーワードを取得（3行目から開始）"""
    try:
        print(f"📊 Google Sheetsからデータを取得中...")
        response = requests.get(CSV_URL)
        response.encoding = 'utf-8'  # UTF-8エンコーディングを明示
        
        # CSVデータを適切に解析
        csv_reader = csv.reader(io.StringIO(response.text))
        keywords = []
        
        for i, row in enumerate(csv_reader):
            # 1行目（ヘッダー）と2行目をスキップして、3行目から開始
            if i <= 1:  
                continue
            
            if len(row) >= 5:  # E列（インデックス4）が存在
                keyword = row[4].strip()
                # 空白文字や特殊文字を除外
                if keyword and keyword != '　' and len(keyword.strip()) > 0:
                    keywords.append(keyword)
        
        print(f"✅ {len(keywords)}個のキーワードを取得しました（3行目から）")
        for i, kw in enumerate(keywords, 1):
            print(f"  {i}. {kw}")
        
        return keywords
        
    except Exception as e:
        print(f"❌ Sheets取得エラー: {e}")
        # フォールバック用のキーワード
        return ["メガリス", "レカネマブ", "タダリス", "埋没法"]

def ensure_dir(p): 
    os.makedirs(p, exist_ok=True)

def make_driver():
    """iPhone 16設定でChromeドライバーを作成"""
    opts = webdriver.ChromeOptions()
    
    # iPhone 16設定
    opts.add_argument("--force-device-scale-factor=1")
    opts.add_argument("--high-dpi-support=1")
    opts.add_argument(f"--window-size={VIEWPORT_W},{VIEWPORT_H + 100}")  # ツールバー分を追加
    
    # モバイルエミュレーション
    mobile_emulation = {
        "deviceMetrics": {"width": VIEWPORT_W, "height": VIEWPORT_H, "pixelRatio": 3.0},
        "userAgent": (
            "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) "
            "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 "
            "Mobile/15E148 Safari/604.1"
        )
    }
    opts.add_experimental_option("mobileEmulation", mobile_emulation)
    
    opts.add_argument("--lang=ja-JP")
    
    if USE_INCOGNITO:
        opts.add_argument("--incognito")
    
    # 自動化検出回避
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    
    # ログレベルを下げる
    opts.add_argument("--log-level=3")
    
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=opts
    )
    
    # webdriver検出を回避
    driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
        "source": """
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
            })
        """
    })
    
    return driver

def find_search_box(driver):
    """Yahoo!モバイルの検索ボックスを確実に見つける"""
    wait = WebDriverWait(driver, 10)
    
    # Yahoo!モバイルの検索ボックスセレクタ
    selectors = [
        "input[name='p']",  # Yahoo!の標準的な検索ボックス
        "input#srchtxt",
        "input[type='search']",
        "input.SearchBox__searchInput",  # モバイル版の可能性
    ]
    
    for selector in selectors:
        try:
            element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
            if element.is_displayed() and element.is_enabled():
                print(f"✅ 検索ボックス発見: {selector}")
                return element
        except:
            continue
    
    raise RuntimeError("検索ボックスが見つかりませんでした")

def clear_and_type(driver, search_box, keyword):
    """検索ボックスを完全にクリアして新しいキーワードを入力"""
    try:
        # フォーカスを当てる
        search_box.click()
        time.sleep(0.2)
        
        # JavaScriptで強制的にクリア
        driver.execute_script("""
            arguments[0].value = '';
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """, search_box)
        time.sleep(0.2)
        
        # 念のため全選択して削除
        search_box.send_keys(Keys.CONTROL + "a")
        time.sleep(0.1)
        search_box.send_keys(Keys.DELETE)
        time.sleep(0.2)
        
        # 文字を1文字ずつ入力
        for i, char in enumerate(keyword):
            search_box.send_keys(char)
            time.sleep(TYPE_INTERVAL)
            
            # 6文字目以降でサジェストチェック（最低6文字は入力）
            if i >= 5:  # 6文字目以降
                if check_yahoo_suggestions(driver):
                    return True, keyword[:i+1]
        
        # 全文字入力後の最終チェック
        time.sleep(PAUSE_AFTER_TYPE)
        if check_yahoo_suggestions(driver):
            return True, keyword
        
        return False, keyword
        
    except Exception as e:
        print(f"❌ 入力エラー: {e}")
        return False, keyword

def check_yahoo_suggestions(driver):
    """Yahoo!のサジェスト表示を確認（高速チェック）"""
    # Yahoo!モバイルの実際のサジェスト要素
    suggest_selectors = [
        # Yahoo!の主要なサジェストセレクタ
        ".sw-SuggestList",
        ".sw-Card",
        ".sw-CardBase",
        ".Suggest",
        "ul.SuggestList",
        
        # より一般的なセレクタ
        "[role='listbox']",
        "div[class*='suggest']",
        "ul[class*='suggest']",
        
        # リストアイテムベース
        "ul li a[href*='search.yahoo']",
    ]
    
    for selector in suggest_selectors:
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)
            for elem in elements:
                # 要素が表示されていて、適切なサイズがある
                if (elem.is_displayed() and 
                    elem.size.get('height', 0) > 20 and 
                    elem.size.get('width', 0) > 100):
                    
                    # 位置が検索ボックスの下にある
                    if elem.location.get('y', 0) > 50:
                        print(f"  🎯 サジェスト検出: {selector}")
                        return True
        except:
            continue
    
    # JavaScriptでより広範囲にチェック
    try:
        result = driver.execute_script("""
            const lists = document.querySelectorAll('ul, ol, div[class*="list"]');
            for (let list of lists) {
                const rect = list.getBoundingClientRect();
                if (rect.height > 50 && rect.top > 100 && rect.top < 500) {
                    const items = list.querySelectorAll('li, a');
                    if (items.length >= 2) {
                        return true;
                    }
                }
            }
            return false;
        """)
        if result:
            print("  🎯 サジェスト検出: JavaScript")
            return True
    except:
        pass
    
    return False

def capture_suggestion(driver, keyword, status="suggest"):
    """スクリーンショットを撮影"""
    # スクロール位置を最上部に
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(SCREENSHOT_DELAY)
    
    # ファイル名を作成
    today = datetime.now().strftime("%Y%m%d")
    timestamp = datetime.now().strftime("%H%M%S")
    safe_keyword = "".join(c for c in keyword if c.isalnum() or c in "ぁ-んァ-ン一-龥")
    filename = f"{today}_{safe_keyword}_{status}_{timestamp}.png"
    filepath = os.path.join(OUTPUT_DIR, filename)
    
    # スクリーンショット撮影
    driver.save_screenshot(filepath)
    print(f"📸 撮影完了: {filename}")
    
    return filepath

def main():
    """メイン処理"""
    ensure_dir(OUTPUT_DIR)
    
    # Google Sheetsからキーワード取得（3行目から）
    keywords = get_keywords_from_sheets()
    
    if not keywords:
        print("❌ キーワードが取得できませんでした")
        return
    
    driver = make_driver()
    results = []
    
    try:
        for i, keyword in enumerate(keywords, 1):
            print(f"\n{'='*50}")
            print(f"📝 {i}/{len(keywords)}: {keyword}")
            print(f"{'='*50}")
            
            try:
                # Yahoo!トップページにアクセス
                driver.get(YAHOO_URL)
                time.sleep(2)
                
                # 検索ボックスを見つける
                search_box = find_search_box(driver)
                
                # キーワードを入力してサジェストチェック
                has_suggest, typed_text = clear_and_type(driver, search_box, keyword)
                
                if has_suggest:
                    filepath = capture_suggestion(driver, keyword, "suggest")
                    results.append(f"✅ {keyword}: サジェスト撮影成功")
                else:
                    filepath = capture_suggestion(driver, keyword, "no_suggest")
                    results.append(f"⚠️ {keyword}: サジェストなし")
                
                # 次のキーワードの前に少し待機
                time.sleep(1)
                
            except Exception as e:
                print(f"❌ エラー: {e}")
                results.append(f"❌ {keyword}: エラー発生")
                continue
        
        # 結果サマリー
        print(f"\n{'='*50}")
        print("📊 処理結果サマリー")
        print(f"{'='*50}")
        for result in results:
            print(result)
        print(f"\n🎉 処理完了！ {len(keywords)}個のキーワードを処理しました")
        
    finally:
        driver.quit()

if __name__ == "__main__":
    main()