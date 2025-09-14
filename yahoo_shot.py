# -*- coding: utf-8 -*-
"""
Yahoo! 検索サジェスト自動撮影ツール（携帯表示・完全版）
- Google Sheets のE列(3行目～)からキーワード取得
- iPhone相当（UA/画面サイズ 393x852）で携帯表示を撮影
- ローカル: 表示あり / CI(GitHub Actions): ヘッドレス
- 競合回避のため専用 user-data-dir を使用
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
from webdriver_manager.chrome import ChromeDriverManager

# ========= 設定 =========
SPREADSHEET_ID = "1i2SKztLstWUeD0y9mNsNtdYrwdtZ7zCzNDA7vhojA8s"
SHEET_GID = "1060586181"
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/export?format=csv&gid={SHEET_GID}"

# 出力先：環境変数があればそれを優先（Actions は /tmp/yahooshot を渡す）
OUTPUT_DIR = os.getenv("YAHOOSHOT_OUTPUT_DIR", r"C:\Users\kanek\Pictures\朝")

# 携帯版のYahooトップ
YAHOO_URL = "https://m.yahoo.co.jp/"

# iPhone 16 相当の画面サイズ（サジェストが1画面に収まる）
VIEWPORT_W, VIEWPORT_H = 393, 852
USE_INCOGNITO = True

# 入力・待機のチューニング
TYPE_INTERVAL = 0.30     # 一文字ごとの待ち
PAUSE_AFTER_TYPE = 1.2   # 全文字入力後の待ち
PAUSE_AFTER_FORCE = 0.6  # サジェスト強制操作後の待ち
SCREENSHOT_DELAY = 0.5   # 撮影前の微待機
# =======================


def ensure_dir(p: str):
    os.makedirs(p, exist_ok=True)


def get_keywords_from_sheets():
    """Google SheetsからE列のキーワード取得（3行目から）。失敗時は固定配列。"""
    try:
        print("Google Sheetsからデータ取得中...")
        r = requests.get(CSV_URL, timeout=15)
        r.encoding = "utf-8"
        r.raise_for_status()
        kws = []
        for i, row in enumerate(csv.reader(io.StringIO(r.text))):
            if i <= 1:  # 1行目ヘッダ＋2行目をスキップ → 3行目から
                continue
            if len(row) >= 5:
                kw = row[4].strip()
                if kw:
                    kws.append(kw)
        if not kws:
            print("E列が空のためフォールバックに切り替え")
            return ["メガリス", "レカネマブ", "タダリス", "埋没法"]
        print(f"取得キーワード数: {len(kws)}")
        for idx, kw in enumerate(kws, 1):
            print(f"  {idx}. {kw}")
        return kws
    except Exception as e:
        print(f"Sheets取得エラー: {e}")
        return ["メガリス", "レカネマブ", "タダリス", "埋没法"]


def make_driver():
    """iPhone設定でChrome起動（ローカル=可視 / Actions=ヘッドレス＋専用プロファイル）"""
    opts = webdriver.ChromeOptions()

    # iPhone相当の見た目
    opts.add_argument("--force-device-scale-factor=1")
    opts.add_argument("--high-dpi-support=1")
    opts.add_argument(f"--window-size={VIEWPORT_W},{VIEWPORT_H + 100}")  # ツールバーぶん少し足す
    mobile_emulation = {
        "deviceMetrics": {"width": VIEWPORT_W, "height": VIEWPORT_H, "pixelRatio": 3.0},
        "userAgent": (
            "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) "
            "AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 "
            "Mobile/15E148 Safari/604.1"
        ),
    }
    opts.add_experimental_option("mobileEmulation", mobile_emulation)
    opts.add_argument("--lang=ja-JP")
    if USE_INCOGNITO:
        opts.add_argument("--incognito")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument("--log-level=3")

    if os.getenv("GITHUB_ACTIONS") == "true":
        # GitHub Actions（Linux, 画面なし）
        opts.add_argument("--headless=new")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        # 実行ごとに一意のプロファイルで衝突回避
        run_id = os.getenv("GITHUB_RUN_ID", str(int(time.time())))
        opts.add_argument(f"--user-data-dir=/tmp/chrome-profile-{run_id}")
    else:
        # ローカルPC：可視・専用プロファイル（競合回避）
        opts.add_argument(r"--user-data-dir=C:/YahooShot/chrome-profile")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=opts
    )
    try:
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {"source": "Object.defineProperty(navigator,'webdriver',{get:()=>undefined});"},
        )
    except Exception:
        pass
    return driver


def find_search_box(driver):
    """携帯版Yahooの検索ボックスを見つける。"""
    wait = WebDriverWait(driver, 10)
    selectors = [
        "input[name='p']",
        "input#srchtxt",
        "input[type='search']",
        "input.SearchBox__searchInput",
    ]
    for css in selectors:
        try:
            el = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, css)))
            if el.is_displayed() and el.is_enabled():
                print(f"検索ボックス: {css}")
                return el
        except Exception:
            continue
    raise RuntimeError("検索ボックスが見つかりませんでした")


def type_slowly(box, text: str, interval: float):
    for ch in text:
        box.send_keys(ch)
        time.sleep(interval)


def capture_viewport(driver, keyword: str, suffix: str = "view"):
    """今見えているビューポートのみ保存（携帯の画面そのまま）。"""
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(SCREENSHOT_DELAY)
    today = datetime.now().strftime("%Y%m%d")
    safe_kw = "".join(c for c in keyword if c not in '\\/:*?"<>|')
    fname = f"{today}_{safe_kw}_{suffix}.png"
    path = os.path.join(OUTPUT_DIR, fname)
    driver.save_screenshot(path)
    print(f"Saved: {path}")
    return path


def main():
    ensure_dir(OUTPUT_DIR)
    keywords = get_keywords_from_sheets()
    if not keywords:
        print("キーワードが取得できませんでした")
        return

    driver = make_driver()
    try:
        for i, kw in enumerate(keywords, 1):
            print("=" * 60)
            print(f"[{i}/{len(keywords)}] {kw}")
            print("=" * 60)

            driver.get(YAHOO_URL)
            time.sleep(2.0)

            box = find_search_box(driver)
            box.click()

            # 1文字ずつ入力 → 少し待つ → サジェスト開かせる
            type_slowly(box, kw, TYPE_INTERVAL)
            time.sleep(PAUSE_AFTER_TYPE)
            box.send_keys(Keys.SPACE); time.sleep(0.1)
            box.send_keys(Keys.BACK_SPACE); time.sleep(0.1)
            box.send_keys(Keys.ARROW_DOWN); time.sleep(PAUSE_AFTER_FORCE)

            capture_viewport(driver, kw, "view")
            time.sleep(0.8)

        print("処理完了")
    finally:
        driver.quit()


if __name__ == "__main__":
    main()
