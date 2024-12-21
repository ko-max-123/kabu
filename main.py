import requests
from bs4 import BeautifulSoup
import os
from openpyxl import Workbook, load_workbook
from datetime import datetime
import time
from plyer import notification
import gc
import threading
import tkinter as tk
from tkinter import ttk
import webbrowser
import textwrap
import re

# ----- Selenium関連 -----
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

# ===== 設定項目 =====
CHECK_URLS = [
    "https://kabutan.jp/news/marketnews/?category=3",
    "https://kabutan.jp/news/marketnews/?category=10",
    "https://kabutan.jp/news/marketnews/?category=2",
    "https://kabutan.jp/news/marketnews/?category=8",
    "https://kabutan.jp/news/marketnews/?category=9"
]
# URLに応じて前回記録用のファイルを振り分け
url_to_lastfile = {}
for url in CHECK_URLS:
    cat = url.split("category=")[-1]
    url_to_lastfile[url] = f"last_article_{cat}.txt"

EXCEL_FILE = "articles.xlsx"                # 記録用Excelファイル
RESULT_FONT = ("TkDefaultFont", 13)         # 検索結果表示用フォント
BODY_FONT = ("TkDefaultFont", 13)           # 別ウィンドウ本文表示用フォント

stop_flag = False
interval_minutes = 1  # デフォルト値（分）
worker_thread = None

# 銘柄コード抽出用の正規表現：<1234> のような4桁数字を取り出す
CODE_PATTERN = re.compile(r"<(\d{4})>")
# 「赤字」「黒字」を色分けするための正規表現
COLOR_PATTERN = re.compile("(赤字|黒字)")

def get_latest_news(url):
    """
    指定URLから最初のニュース(最上部tr)の情報を取得
    (ニュース時刻, カテゴリ, タイトル, フルURL) を返す
    """
    response = requests.get(url)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # 修正箇所： tbody → 削除
    tr = soup.select_one("table.s_news_list tr")
    if not tr:
        del soup, response
        return None
    
    time_el = tr.select_one("td.news_time time")
    news_time = time_el.get_text(strip=True) if time_el else ""
    
    category_div = tr.select_one("td > div.newslist_ctg")
    category = category_div.get_text(strip=True) if category_div else ""
    
    link = tr.select_one("td > a")
    title = link.get_text(strip=True) if link else ""
    href = link.get("href", "") if link else ""
    full_url = "https://kabutan.jp" + href if href.startswith("/") else href

    del soup, response, tr, time_el, category_div, link
    if title:
        return (news_time, category, title, full_url)
    return None

def read_last_record(file_path):
    if os.path.exists(file_path):
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read().strip()
    return None

def write_last_record(file_path, record):
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(record)

def append_to_excel(file_path, news):
    news_time, category, title, url = news

    if os.path.exists(file_path):
        wb = load_workbook(file_path)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["DateTimeChecked", "NewsTime", "Category", "Title", "URL"])
        
    ws.append([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        news_time,
        category,
        title,
        url
    ])
    wb.save(file_path)
    del wb, ws

def show_notification(title, message):
    notification.notify(
        title=title,
        message=message,
        timeout=5
    )

def open_sbi_with_code(code_list):
    """
    SBI証券サイトを開き、account.txtに記載のID/PWでログインし、
    その後、株価検索フォームに銘柄コードを入力して検索する。
    ブラウザを閉じず、Seleniumの制御も解除
    """
    if not code_list:
        print("銘柄コードがありません。操作を中断します。")
        return

    code = code_list[0]  # 例として先頭の銘柄コードを検索
    sbi_url = "https://site2.sbisec.co.jp/ETGate/"

    if not os.path.exists("account.txt"):
        print("account.txt が存在しません。ID/PWを設定してください。")
        return
    with open("account.txt", "r", encoding="utf-8") as f:
        lines = f.read().splitlines()
    if len(lines) < 2:
        print("account.txt にIDとPWを正しく設定してください。(1行目:ID, 2行目:PW)")
        return
    user_id = lines[0].strip()
    user_pw = lines[1].strip()

    chrome_options = Options()
    # detachにより、スクリプト終了後もブラウザを閉じない
    chrome_options.add_experimental_option("detach", True)
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--start-maximized")

    driver = webdriver.Chrome(options=chrome_options)
    try:
        driver.get(sbi_url)
        time.sleep(2)

        # ログインフォーム要素名を適宜修正
        user_box = driver.find_element(By.NAME, "user_id")
        pass_box = driver.find_element(By.NAME, "user_password")
        user_box.send_keys(user_id)
        pass_box.send_keys(user_pw)
        login_btn = driver.find_element(By.NAME, "ACT_login")
        login_btn.click()

        time.sleep(3)  # ログイン処理待ち

        try:
            # 修正箇所: stock_search_box の要素名
            stock_search_box = driver.find_element(By.NAME, "i_stock_sec")
            stock_search_box.send_keys(code)
            # 修正箇所: 検索ボタン -> CSS_SELECTOR "#srchK > a"
            driver.find_element(By.CSS_SELECTOR, "#srchK > a").click()
        except Exception as e:
            print("株価検索フォームが見つからない、または操作失敗:", e)
            return

        print(f"銘柄コード {code} を検索しました。")
    except Exception as e:
        print("SBI証券での操作中にエラーが発生:", e)

def show_body_window(url):
    """
    記事本文を別ウィンドウで表示
    (銘柄コード抽出/「証券会社のページを開く」ボタン等)
    - brタグを改行に置き換え
    - 「。」で改行
    """
    try:
        resp = requests.get(url)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, 'html.parser')
        
        body_elem = soup.select_one(".body")
        if body_elem:
            # brタグを改行に置き換え
            for br in body_elem.find_all("br"):
                br.replace_with("\n")

            body_text = body_elem.get_text()
            # 「。」があるときは改行を挿入
            body_text = body_text.replace("。", "。\n")
        else:
            body_text = "本文を取得できませんでした。"
        del soup, resp
    except Exception as e:
        body_text = f"エラーが発生しました: {e}"

    win = tk.Toplevel(root)
    win.title("記事本文")

    content_frame = ttk.Frame(win, padding=10)
    content_frame.grid(sticky=(tk.W, tk.E, tk.N, tk.S))

    canvas = tk.Canvas(content_frame, width=800, height=600)
    canvas.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

    scrollbar = ttk.Scrollbar(content_frame, orient=tk.VERTICAL, command=canvas.yview)
    scrollbar.grid(row=0, column=3, sticky=(tk.N, tk.S))
    canvas.configure(yscrollcommand=scrollbar.set)

    body_frame = ttk.Frame(canvas)
    canvas.create_window((0,0), window=body_frame, anchor='nw')

    # body_text を行ごとに分割して表示 (改行が増えたぶん分割しやすい)
    lines = body_text.split("\n")
    for line in lines:
        line_label = tk.Label(body_frame, text=line, font=BODY_FONT, justify="left")
        line_label.pack(anchor='w', pady=2, padx=5)

    def on_configure(event):
        canvas.config(scrollregion=canvas.bbox("all"))
    body_frame.bind("<Configure>", on_configure)

    # 銘柄コード
    codes = CODE_PATTERN.findall(body_text)
    if codes:
        codes_label = tk.Label(content_frame, text=f"取得した銘柄コード: {', '.join(codes)}", fg="green", font=BODY_FONT)
        codes_label.grid(row=1, column=0, columnspan=3, pady=5, sticky=tk.W)
    else:
        codes_label = tk.Label(content_frame, text="銘柄コードは検出されませんでした", fg="gray", font=BODY_FONT)
        codes_label.grid(row=1, column=0, columnspan=3, pady=5, sticky=tk.W)

    def open_original():
        webbrowser.open(url)
    open_button = tk.Button(content_frame, text="元ページを開く", command=open_original, bg="blue", fg="white", font=BODY_FONT)
    open_button.grid(row=2, column=0, pady=10, sticky=tk.W)

    def open_sbi():
        open_sbi_with_code(codes)
    sbi_button = tk.Button(content_frame, text="証券会社のページを開く", command=open_sbi, bg="green", fg="white", font=BODY_FONT)
    sbi_button.grid(row=2, column=1, pady=10, padx=10, sticky=tk.W)

    win.columnconfigure(0, weight=1)
    win.rowconfigure(0, weight=1)
    content_frame.columnconfigure(0, weight=1)
    content_frame.rowconfigure(0, weight=1)

def format_result_text(parent, text):
    """
    検索結果の1行テキストから「赤字」を赤、「黒字」を青で表示
    """
    pos = 0
    matches = list(COLOR_PATTERN.finditer(text))
    for m in matches:
        keyword = m.group(0)
        color = "red" if keyword == "赤字" else "blue"
        start = m.start()
        if start > pos:
            segment = text[pos:start]
            lbl = tk.Label(parent, text=segment, font=RESULT_FONT)
            lbl.pack(side='left', padx=0)
        seg_lbl = tk.Label(parent, text=keyword, font=RESULT_FONT, fg=color)
        seg_lbl.pack(side='left', padx=0)
        pos = m.end()
    if pos < len(text):
        remainder = text[pos:]
        lbl = tk.Label(parent, text=remainder, font=RESULT_FONT)
        lbl.pack(side='left', padx=0)

def print_result(msg, url=None):
    """
    結果表示用エリアに1行追加。
    URLがあれば「記事を見る」ボタンを付加。
    """
    def insert_line():
        line_frame = ttk.Frame(results_frame)
        line_frame.pack(fill='x', pady=2, anchor='w')
        
        format_result_text(line_frame, msg)

        if url:
            def on_click():
                show_body_window(url)
            btn = tk.Button(line_frame, text="記事を見る", command=on_click, fg="blue", font=RESULT_FONT)
            btn.pack(side='left', padx=5)

        results_frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))

    results_frame.after(0, insert_line)

def check_for_update_on_url(url):
    """
    1つのURLに対してニュース更新チェックを行う。
    """
    latest_news = get_latest_news(url)
    if not latest_news:
        print_result(f"{datetime.now()}: {url} 記事を取得できませんでした。")
        return
    
    news_time, category, latest_title, full_url = latest_news

    last_file = url_to_lastfile[url]
    last_record = read_last_record(last_file)
    
    if latest_title != last_record:
        append_to_excel(EXCEL_FILE, (news_time, category, latest_title, full_url))
        show_notification("新着記事", f"新しい記事が更新されました: {latest_title}")
        print_result(f"{news_time}: [{category}] {latest_title}", url=full_url)
        write_last_record(last_file, latest_title)
    #else:
    #    print_result(f"{datetime.now()}: [{url}] 更新がありません")

    del latest_news, last_record, latest_title, full_url
    gc.collect()

def check_for_update_all_urls():
    """
    CHECK_URLSに含まれる複数URLを一括チェック
    """
    for url in CHECK_URLS:
        check_for_update_on_url(url)

def scraping_worker():
    global stop_flag, interval_minutes
    while not stop_flag:
        check_for_update_all_urls()
        for _ in range(interval_minutes * 60):
            if stop_flag:
                break
            time.sleep(1)
    print_result("スクレイピング停止")

def start_scraping():
    global stop_flag, worker_thread, interval_minutes
    try:
        val = interval_entry.get().strip()
        if val:
            interval_minutes = int(val)
        else:
            interval_minutes = 1
    except ValueError:
        interval_minutes = 1

    stop_flag = False
    if worker_thread is None or not worker_thread.is_alive():
        worker_thread = threading.Thread(target=scraping_worker, daemon=True)
        worker_thread.start()
        print_result(f"スクレイピング開始 (間隔: {interval_minutes}分)")

def stop_scraping():
    global stop_flag
    stop_flag = True
    print_result("停止ボタンが押されました。")

# ==================== GUI構築 ====================
root = tk.Tk()
root.title("ニューススクレイパー (複数ページ対応)")

frame = ttk.Frame(root, padding=20)
frame.grid(sticky=(tk.W, tk.E, tk.N, tk.S))

interval_label = ttk.Label(frame, text="間隔(分):")
interval_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)

interval_entry = ttk.Entry(frame, width=10)
interval_entry.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
interval_entry.insert(0, "1")

start_button = tk.Button(frame, text="実行", bg="blue", fg="white",
                         command=start_scraping, width=10, font=RESULT_FONT)
start_button.grid(row=0, column=2, padx=5, pady=5, sticky=(tk.W))

stop_button = tk.Button(frame, text="停止", bg="red", fg="white",
                        command=stop_scraping, width=10, font=RESULT_FONT)
stop_button.grid(row=0, column=3, padx=5, pady=5, sticky=(tk.W))

canvas = tk.Canvas(frame, width=800, height=400)
canvas.grid(row=1, column=0, columnspan=4, padx=5, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=canvas.yview)
scrollbar.grid(row=1, column=4, sticky=(tk.N, tk.S))
canvas.configure(yscrollcommand=scrollbar.set)

results_frame = ttk.Frame(canvas)
canvas.create_window((0,0), window=results_frame, anchor='nw')

def on_configure(event):
    canvas.config(scrollregion=canvas.bbox("all"))

results_frame.bind("<Configure>", on_configure)

root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
frame.columnconfigure(1, weight=1)
frame.rowconfigure(1, weight=1)

root.mainloop()
