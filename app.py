import io, re, time, traceback, json, os, threading, datetime, schedule, sys, sqlite3, uuid
import pandas as pd
from flask import Flask, render_template_string, request, send_file, redirect, jsonify
from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.styles import PatternFill

app = Flask(__name__)
app.secret_key = 'supersecretkey'
BASE_URL = "https://www.mvideo.ru"
MSK_TZ = datetime.timezone(datetime.timedelta(hours=3))

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
os.makedirs(DATA_DIR, exist_ok=True)
DOWNLOAD_DIR = os.path.join(BASE_DIR, 'downloads')
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
DB_FILE = os.path.join(DATA_DIR, 'mvideo.db')

# Хранилище задач (thread-safe)
tasks = {}
tasks_lock = threading.Lock()

previous_result = []
previous_file = 'last_result.json'

if os.path.exists(previous_file):
    try:
        with open(previous_file, 'r', encoding='utf-8') as f:
            previous_result = json.load(f)
    except:
        previous_result = []

# ===== HTML-шаблон (обновлён с поддержкой task_id) =====
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Парсер для маркетплейса М.Видео</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            background: linear-gradient(135deg, #8B0000 0%, #4A0000 100%);
            display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0;
        }
        .container {
            background: #FFFFFF; padding: 2.5rem; border-radius: 16px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.3); width: 100%; max-width: 780px; text-align: left;
        }
        h1 { font-size: 1.8rem; font-weight: 600; color: #1E1E1E; margin-bottom: 0.2rem; letter-spacing: -0.3px; }
        .version-info {
            font-size: 0.75rem; color: #888; margin-bottom: 1.5rem; line-height: 1.3;
        }
        .input-area { position: relative; margin-bottom: 1.5rem; }
        textarea {
            width: 100%; height: 120px; padding: 1rem; border: 1px solid #E0E0E0;
            border-radius: 12px; font-size: 0.95rem; line-height: 1.4; resize: vertical;
            background: #FAFAFA; transition: border 0.2s; font-family: inherit;
        }
        textarea:focus { outline: none; border-color: #D32F2F; background: #FFF; }
        .send-btn {
            position: absolute; bottom: 12px; right: 12px; background: #D32F2F; color: white;
            border: none; width: 40px; height: 40px; border-radius: 50%; cursor: pointer;
            display: flex; align-items: center; justify-content: center;
            transition: background 0.2s, transform 0.1s; box-shadow: 0 2px 6px rgba(211,47,47,0.3);
            font-size: 1.2rem;
        }
        .send-btn:hover { background: #B71C1C; }
        .send-btn:disabled { background: #BDBDBD; cursor: not-allowed; box-shadow: none; }
        .send-btn:active { transform: scale(0.95); }
        .flash-messages { list-style: none; padding: 0; margin: 1rem 0; }
        .flash-messages li { padding: 0.6rem 1rem; border-radius: 8px; margin-bottom: 0.5rem; font-size: 0.9rem; }
        .success { background: #E8F5E9; color: #2E7D32; }
        .error { background: #FFEBEE; color: #C62828; }
        #loading-indicator {
            display: none; text-align: center; margin: 1.5rem 0;
        }
        .spinner {
            width: 32px; height: 32px; margin: 0 auto 0.8rem;
            border: 3px solid #F0F0F0; border-top-color: #D32F2F;
            border-radius: 50%; animation: spin 0.8s linear infinite;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
        .loading-text { color: #666; font-size: 0.9rem; font-weight: 500; }
        .progress-text { color: #333; font-size: 0.9rem; font-weight: 500; margin-top: 0.3rem; }
        #actions { margin-top: 1.5rem; }
        .btn-group { display: flex; flex-wrap: wrap; gap: 0.8rem; margin-bottom: 1.5rem; }
        .btn {
            background: #FFF; color: #D32F2F; border: 1px solid #D32F2F;
            padding: 0.6rem 1.4rem; border-radius: 8px; font-size: 0.9rem; cursor: pointer;
            transition: all 0.2s; text-decoration: none; display: inline-block; font-weight: 500;
        }
        .btn:hover { background: #FFEBEE; }
        .btn-primary { background: #D32F2F; color: #FFF; border: none; }
        .btn-primary:hover { background: #B71C1C; }
        .seller-block {
            display: flex; align-items: center; gap: 0.5rem; margin-top: 1.5rem; border-top: 1px solid #F0F0F0; padding-top: 1.5rem;
        }
        .seller-block label { font-weight: 500; color: #444; white-space: nowrap; font-size: 0.9rem; }
        .seller-block input {
            flex: 1; padding: 0.5rem 0.8rem; border: 1px solid #E0E0E0; border-radius: 8px; font-size: 0.9rem;
            background: #FAFAFA; transition: border 0.2s;
        }
        .seller-block input:focus { outline: none; border-color: #D32F2F; background: #FFF; }
        .seller-block .btn { white-space: nowrap; }
        .history-block {
            margin-top: 2rem; border-top: 1px solid #F0F0F0; padding-top: 1.5rem;
        }
        .history-header {
            display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.8rem;
        }
        .history-header h3 { font-size: 1.1rem; font-weight: 600; color: #333; }
        .history-search {
            padding: 0.4rem 0.8rem; border: 1px solid #E0E0E0; border-radius: 8px;
            font-size: 0.85rem; background: #FAFAFA; width: 220px; transition: border 0.2s;
        }
        .history-search:focus { outline: none; border-color: #D32F2F; background: #FFF; }
        .history-list { list-style: none; padding: 0; margin: 0; max-height: 280px; overflow-y: auto; }
        .history-list li {
            display: flex; justify-content: space-between; align-items: center;
            padding: 0.5rem 0; border-bottom: 1px solid #F5F5F5; font-size: 0.9rem;
        }
        .history-list li:last-child { border-bottom: none; }
        .history-list a { color: #D32F2F; text-decoration: none; font-weight: 500; }
        .history-list a:hover { text-decoration: underline; }
        .stats {
            margin-top: 1rem; font-size: 0.8rem; color: #888; text-align: right;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Парсер для маркетплейса М.Видео</h1>
        <p class="version-info">Версия 1.0.0 — умеет собирать код товара, цену, ссылки на товары и сравнивать цены; при отклонении 5% и более выгружает список товаров.</p>

        <ul class="flash-messages" id="flash-messages"></ul>

        <div class="input-area">
            <textarea id="urls" name="urls" placeholder="https://www.mvideo.ru/seller/K000071378" onkeydown="if(event.key==='Enter' && event.shiftKey){ event.preventDefault(); startParsing(); }">{{ request.form.get('urls', '') }}</textarea>
            <button class="send-btn" id="start-btn" onclick="startParsing()" title="Запустить парсинг (Shift+Enter)">&#9654;</button>
        </div>

        <div id="loading-indicator">
            <div class="spinner"></div>
            <p class="loading-text">Сбор данных, пожалуйста, подождите...</p>
            <p class="progress-text" id="progress-count" style="display:none;">Собрано данных с 0 товаров</p>
        </div>

        <div id="actions">
            <div class="btn-group">
                <a href="/download" class="btn btn-primary" id="download-link">Скачать Excel</a>
                <button class="btn btn-primary" onclick="downloadDiscrepancies()">Выгрузить расхождения цен (≥5%)</button>
            </div>
            <div class="seller-block">
                <label for="seller-search">Продавец для истории:</label>
                <input type="text" id="seller-search" list="sellers-list" placeholder="Введите или выберите ссылку">
                <datalist id="sellers-list">
                    {% for s in sellers %}
                    <option value="{{ s }}">
                    {% endfor %}
                </datalist>
                <button class="btn" onclick="downloadHistory()">История цен за неделю</button>
            </div>

            <div class="history-block">
                <div class="history-header">
                    <h3>Последние запросы</h3>
                    <input type="text" id="history-search-input" class="history-search" placeholder="Поиск по истории..." oninput="filterHistory()">
                </div>
                <ul class="history-list" id="history-list">
                    {% for item in history %}
                    <li>
                        <span>{{ item.timestamp }} — {{ item.product_count }} товаров</span>
                        <a href="/download_history/{{ item.filename }}">Скачать</a>
                    </li>
                    {% endfor %}
                    {% if not history %}
                    <li>Нет сохранённых запросов.</li>
                    {% endif %}
                </ul>
            </div>
            <div class="stats">
                Обработано за всё время: {{ total_all_time }} товаров | За последние 24 часа: {{ total_last_24h }} товаров
            </div>
        </div>
    </div>

    <script>
        let currentTaskId = null;
        let statusInterval = null;

        function showMessage(category, text) {
            const container = document.getElementById('flash-messages');
            const li = document.createElement('li');
            li.className = category;
            li.textContent = text;
            container.appendChild(li);
        }

        function startParsing() {
            document.getElementById('start-btn').disabled = true;
            document.getElementById('actions').style.display = 'none';
            document.getElementById('loading-indicator').style.display = 'block';
            document.getElementById('progress-count').style.display = 'none';
            document.getElementById('flash-messages').innerHTML = '';

            const urls = document.getElementById('urls').value;
            fetch('/', {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: 'urls=' + encodeURIComponent(urls)
            })
            .then(response => response.json())
            .then(data => {
                if (data.task_id) {
                    currentTaskId = data.task_id;
                    statusInterval = setInterval(checkStatus, 1000);
                } else {
                    showMessage('error', data.message || 'Ошибка запуска');
                    resetUI();
                }
            });
        }

        function checkStatus() {
            if (!currentTaskId) return;
            fetch('/status?task_id=' + currentTaskId)
                .then(response => response.json())
                .then(data => {
                    if (data.running) {
                        if (data.completed > 0) {
                            const progress = document.getElementById('progress-count');
                            progress.textContent = 'Собрано данных с ' + data.completed + ' товаров';
                            progress.style.display = 'block';
                        }
                    } else {
                        clearInterval(statusInterval);
                        document.getElementById('start-btn').disabled = false;
                        document.getElementById('loading-indicator').style.display = 'none';
                        document.getElementById('progress-count').style.display = 'none';
                        document.getElementById('actions').style.display = 'block';
                        if (data.success) {
                            showMessage('success', 'Парсинг успешно завершён. Обработано ' + data.completed + ' товаров.');
                            // Обновляем ссылку на скачивание с task_id
                            const downloadLink = document.getElementById('download-link');
                            downloadLink.href = '/download?task_id=' + currentTaskId;
                        } else {
                            showMessage('error', 'Парсинг завершён, но возникли ошибки.');
                        }
                        location.reload();
                    }
                });
        }

        function resetUI() {
            document.getElementById('start-btn').disabled = false;
            document.getElementById('loading-indicator').style.display = 'none';
            document.getElementById('actions').style.display = 'block';
        }

        function downloadDiscrepancies() {
            window.location.href = '/discrepancies';
        }

        function downloadHistory() {
            const seller = document.getElementById('seller-search').value.trim();
            if (seller) {
                window.location.href = '/history?seller_url=' + encodeURIComponent(seller);
            } else {
                showMessage('error', 'Введите или выберите ссылку продавца.');
            }
        }

        function filterHistory() {
            const input = document.getElementById('history-search-input');
            const filter = input.value.toLowerCase();
            const list = document.getElementById('history-list');
            const items = list.getElementsByTagName('li');
            for (let i = 0; i < items.length; i++) {
                const text = items[i].textContent.toLowerCase();
                items[i].style.display = text.includes(filter) ? '' : 'none';
            }
        }
    </script>
</body>
</html>
"""

# ===== База данных SQLite =====
def get_db_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS sellers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            seller_url TEXT NOT NULL,
            code TEXT NOT NULL,
            UNIQUE(seller_url, code)
        );
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS price_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT NOT NULL,
            date TEXT NOT NULL,
            price INTEGER,
            UNIQUE(code, date)
        );
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS parsing_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL,
            filename TEXT NOT NULL,
            product_count INTEGER NOT NULL
        );
    """)
    conn.commit()
    conn.close()

def save_seller_product(seller_url, code):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        INSERT OR IGNORE INTO sellers (seller_url, code)
        VALUES (?, ?)
    """, (seller_url, code))
    conn.commit()
    conn.close()

def save_price(code, price):
    today = datetime.date.today().isoformat()
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO price_history (code, date, price)
        VALUES (?, ?, ?)
        ON CONFLICT(code, date) DO UPDATE SET price = excluded.price;
    """, (code, today, price))
    conn.commit()
    conn.close()

def clean_old_prices():
    cutoff = (datetime.date.today() - datetime.timedelta(days=7)).isoformat()
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("DELETE FROM price_history WHERE date < ?;", (cutoff,))
    conn.commit()
    conn.close()

def get_all_sellers():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT seller_url FROM sellers;")
    rows = [r[0] for r in cur.fetchall()]
    conn.close()
    return rows

def save_parsing_result(result_data, product_count):
    now_moscow = datetime.datetime.now(MSK_TZ)
    timestamp_str = now_moscow.strftime('%d.%m.%Y %H:%M')
    filename = f"parser_{now_moscow.strftime('%d.%m.%Y_%H:%M')}.xlsx"
    if result_data:
        df = pd.DataFrame(result_data)
        filepath = os.path.join(DOWNLOAD_DIR, filename)
        df.to_excel(filepath, index=False, engine='openpyxl')
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("""
            INSERT INTO parsing_history (timestamp, filename, product_count)
            VALUES (?, ?, ?)
        """, (timestamp_str, filename, product_count))
        conn.commit()
        conn.close()

def get_parsing_history(limit=100):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT timestamp, filename, product_count FROM parsing_history
        ORDER BY id DESC LIMIT ?
    """, (limit,))
    rows = cur.fetchall()
    conn.close()
    return rows

def get_total_stats():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT COALESCE(SUM(product_count), 0) FROM parsing_history")
    total_all = cur.fetchone()[0]
    now_moscow = datetime.datetime.now(MSK_TZ)
    last_24h = now_moscow - datetime.timedelta(hours=24)
    last_24h_str = last_24h.strftime('%d.%m.%Y %H:%M')
    cur.execute("""
        SELECT COALESCE(SUM(product_count), 0) FROM parsing_history
        WHERE timestamp >= ?
    """, (last_24h_str,))
    total_24h = cur.fetchone()[0]
    conn.close()
    return total_all, total_24h

# ===== Парсинг (с поддержкой task_id) =====
def log_error(msg):
    print(f"[ERROR] {msg}")

def add_status(task, msg):
    if task:
        with tasks_lock:
            task['messages'].append(msg)

def parse_seller_page_final(context, seller_url, task=None):
    for attempt in range(2):
        page = context.new_page()
        try:
            add_status(task, f"Попытка {attempt+1}: загружаем {seller_url}")
            resp = page.goto(seller_url, wait_until='load', timeout=30000)
            if resp.status >= 400:
                log_error(f"Сайт вернул ошибку {resp.status}: {resp.status_text}")
                page.close()
                return None

            try:
                page.wait_for_selector('a[href*="/products/"]', timeout=20000)
            except:
                add_status(task, "Товары не появились, страница пустая")
                page.close()
                return []

            max_clicks = 400
            clicks = 0
            no_change_streak = 0

            while clicks < max_clicks:
                button = page.query_selector('button:has-text("Показать ещё"), span:has-text("Показать ещё")')
                if not button or not button.is_visible():
                    page.evaluate('window.scrollTo(0, document.body.scrollHeight)')
                    page.wait_for_timeout(2000)
                    button = page.query_selector('button:has-text("Показать ещё"), span:has-text("Показать ещё")')
                    if not button or not button.is_visible():
                        add_status(task, "Кнопка «Показать ещё» больше не найдена")
                        break

                prev_count = page.evaluate('document.querySelectorAll(\'a[href*="/products/"]\').length')
                add_status(task, f"Клик {clicks+1}, товаров сейчас: {prev_count}")

                button.scroll_into_view_if_needed()
                try:
                    button.click()
                except:
                    page.evaluate('(btn) => btn.click()', button)

                try:
                    page.wait_for_function(
                        f'() => document.querySelectorAll(\'a[href*="/products/"]\').length > {prev_count}',
                        timeout=8000
                    )
                    no_change_streak = 0
                except:
                    page.wait_for_timeout(2000)
                    no_change_streak += 1
                    if no_change_streak >= 3:
                        add_status(task, "3 раза количество не изменилось, завершаем загрузку")
                        break

                clicks += 1

            for _ in range(3):
                page.evaluate('window.scrollTo(0, document.body.scrollHeight)')
                page.wait_for_timeout(1500)

            products_data = page.evaluate('''() => {
                const items = [];
                const links = document.querySelectorAll('a[href*="/products/"]');
                links.forEach(link => {
                    const href = link.getAttribute('href');
                    let priceSpan = null;
                    let el = link;
                    for (let i = 0; i < 5 && el; i++) {
                        const priceContainer = el.querySelector('div.price');
                        if (priceContainer) {
                            priceSpan = priceContainer.querySelector('span.current-price');
                            if (priceSpan) break;
                        }
                        el = el.parentElement;
                    }
                    items.push({ href, priceText: priceSpan ? priceSpan.textContent.trim() : null });
                });
                return items;
            }''')

            if not products_data:
                log_error("Не удалось извлечь товары")
                page.close()
                return []

            products = []
            for item in products_data:
                href = item.get('href')
                if not href:
                    continue
                full_url = href if href.startswith('http') else BASE_URL + href
                code_match = re.search(r'(\d+)$', full_url)
                code = code_match.group(1) if code_match else "N/A"
                price = None
                price_text = item.get('priceText')
                if price_text:
                    digits = re.sub(r'\D', '', price_text)
                    if digits:
                        try:
                            price = int(digits)
                        except:
                            pass
                products.append({
                    'productId': code,
                    'price': price if price is not None else "N/A",
                    'link': full_url
                })
            page.close()
            return products
        except Exception as e:
            log_error(f"Ошибка при парсинге магазина: {e}")
            page.close()
        time.sleep(5)
    return None

def parsing_thread(urls, task_id=None, save_to_db=True):
    task = None
    if task_id:
        with tasks_lock:
            task = tasks.get(task_id)
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=[
                    '--disable-blink-features=AutomationControlled',
                    '--no-sandbox',
                    '--disable-web-security',
                    '--disable-features=IsolateOrigins,site-per-process',
                ]
            )
            context = browser.new_context(
                user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36',
                viewport={'width': 1440, 'height': 900},
                locale='ru-RU',
                timezone_id='Europe/Moscow',
            )
            context.add_init_script("""
                Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                window.chrome = { runtime: {} };
                Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]});
                Object.defineProperty(navigator, 'languages', {get: () => ['ru-RU', 'ru']});
            """)
            context.set_extra_http_headers({
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
                'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
                'Referer': 'https://www.mvideo.ru/',
            })
            context.set_extra_http_headers({
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
                'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
                'Referer': 'https://www.mvideo.ru/',
            })
            context.add_cookies([
                {'name': 'MVID_CITY_ID', 'value': 'CityCZ_975', 'domain': '.mvideo.ru', 'path': '/'},
                {'name': 'MVID_REGION_SHOP', 'value': 'S002', 'domain': '.mvideo.ru', 'path': '/'}
            ])

            all_data = []
            for seller_url in urls:
                add_status(task, f"Обрабатывается магазин: {seller_url}")
                products = parse_seller_page_final(context, seller_url, task)
                if products is None:
                    log_error(f"Сайт заблокировал доступ для {seller_url}")
                    continue
                if not products:
                    continue
                if task:
                    with tasks_lock:
                        task['total'] += len(products)
                add_status(task, f"Найдено {len(products)} товаров, начинаем запись")

                if save_to_db:
                    conn = get_db_connection()
                    cur = conn.cursor()
                    cur.execute("DELETE FROM sellers WHERE seller_url = ?;", (seller_url,))
                    conn.commit()
                    conn.close()

                for prod in products:
                    if task:
                        with tasks_lock:
                            task['current_link'] = prod['link']
                            task['completed'] += 1
                    all_data.append({
                        'Код товара': prod['productId'],
                        'Цена': prod['price'] if prod['price'] != "N/A" else "N/A",
                        'Ссылка на товар': prod['link'],
                        'Магазин': seller_url
                    })
                    if save_to_db and prod['productId'] != "N/A":
                        save_seller_product(seller_url, prod['productId'])
                        if isinstance(prod['price'], int):
                            save_price(prod['productId'], prod['price'])
            browser.close()
            # Сохраняем результат в задачу
            if task:
                with tasks_lock:
                    task['result'] = all_data
                    task['success'] = True
            if save_to_db:
                clean_old_prices()
                save_parsing_result(all_data, len(all_data))
            add_status(task, f"Парсинг завершён. Всего товаров: {len(all_data)}")
            # Сохраняем last_result глобально для совместимости со старой кнопкой (пока не нужно)
            global previous_result
            with open(previous_file, 'w', encoding='utf-8') as f:
                json.dump(all_data, f, ensure_ascii=False, indent=2)
            previous_result = all_data
    except Exception as e:
        log_error(f"Глобальная ошибка: {e}")
        traceback.print_exc()
        if task:
            with tasks_lock:
                task['success'] = False
    finally:
        if task:
            with tasks_lock:
                task['running'] = False

# ===== Планировщик (каждые 6 часов) =====
def run_schedule():
    while True:
        schedule.run_pending()
        time.sleep(1)

def scheduled_parse():
    # Планировщик не мешает пользовательским задачам
    sellers = get_all_sellers()
    if sellers:
        print("Автоматический парсинг запущен для", len(sellers), "продавцов.")
        # Запускаем без task_id (просто обновление БД)
        parsing_thread(sellers, task_id=None, save_to_db=True)

# ===== Маршруты =====
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        urls_text = request.form.get('urls', '')
        urls = [u.strip() for u in urls_text.splitlines() if u.strip()]
        if not urls:
            return jsonify({'status': 'error', 'message': 'Введите хотя бы одну ссылку.'})

        # Проверяем, нет ли уже запущенной задачи (можно разрешить несколько)
        # Создаём новую задачу
        task_id = str(uuid.uuid4())
        task = {
            'running': True,
            'total': 0,
            'completed': 0,
            'start_time': time.time(),
            'current_link': '',
            'messages': [],
            'result': [],
            'success': False
        }
        with tasks_lock:
            tasks[task_id] = task

        thread = threading.Thread(target=parsing_thread, args=(urls, task_id, True))
        thread.daemon = True
        thread.start()
        return jsonify({'status': 'started', 'task_id': task_id})

    sellers = get_all_sellers()
    history = get_parsing_history(100)
    total_all, total_24h = get_total_stats()
    return render_template_string(HTML_TEMPLATE, sellers=sellers, history=history,
                                  total_all_time=total_all, total_last_24h=total_24h)

@app.route('/status')
def status():
    task_id = request.args.get('task_id')
    if not task_id:
        return jsonify({'error': 'Missing task_id'}), 400
    with tasks_lock:
        task = tasks.get(task_id)
        if not task:
            return jsonify({'error': 'Task not found'}), 404
        progress = 0
        if task['total'] > 0:
            progress = round((task['completed'] / task['total']) * 100, 1)
        remaining_time = None
        if task['completed'] > 0 and task['running']:
            elapsed = time.time() - task['start_time']
            avg_time = elapsed / task['completed']
            remaining = (task['total'] - task['completed']) * avg_time
            if remaining < 60:
                remaining_time = f"{int(remaining)} сек"
            else:
                remaining_time = f"{int(remaining // 60)} мин {int(remaining % 60)} сек"
        return jsonify({
            'running': task['running'],
            'total': task['total'],
            'completed': task['completed'],
            'percent': progress,
            'current_link': task['current_link'],
            'remaining_time': remaining_time,
            'messages': task['messages'][-5:],
            'success': not task['running'] and task.get('success', False)
        })

@app.route('/download')
def download():
    task_id = request.args.get('task_id')
    if task_id:
        with tasks_lock:
            task = tasks.get(task_id)
            if not task or not task.get('result'):
                return '<p>Нет данных для этой задачи.</p>'
            data = task['result']
    else:
        # fallback: последний глобальный результат
        global previous_result
        data = previous_result
        if not data:
            return '<p>Нет данных для скачивания.</p>'

    now_moscow = datetime.datetime.now(MSK_TZ)
    filename = f"parser_{now_moscow.strftime('%d.%m.%Y_%H:%M')}.xlsx"
    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Товары')
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=filename)

@app.route('/discrepancies')
def discrepancies():
    today = datetime.date.today().isoformat()
    yesterday = (datetime.date.today() - datetime.timedelta(days=1)).isoformat()
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT s.seller_url, p_today.code, p_today.price AS price_today,
               p_yest.price AS price_yesterday
        FROM price_history p_today
        JOIN sellers s ON s.code = p_today.code
        LEFT JOIN price_history p_yest ON p_today.code = p_yest.code AND p_yest.date = ?
        WHERE p_today.date = ?
          AND p_yest.price IS NOT NULL
          AND p_yest.price > 0
          AND ABS((p_today.price - p_yest.price) * 1.0 / p_yest.price) >= 0.05
    """, (yesterday, today))
    rows = cur.fetchall()
    conn.close()

    if not rows:
        return '<p>Нет товаров с изменением цены ≥5% относительно вчерашнего дня.</p>'

    wb = Workbook()
    ws = wb.active
    ws.title = "Расхождения"
    ws.append(["Ссылка на продавца", "Код товара", "Цена сегодня", "Цена вчера", "Изменение, %"])

    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    orange_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
    green_fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")

    for row in rows:
        seller_url, code, price_today, price_yesterday = row
        change = (price_today - price_yesterday) / price_yesterday * 100
        excel_row = [seller_url, code, price_today, price_yesterday, round(change, 1)]
        ws.append(excel_row)
        if change >= 30:
            fill = red_fill
        elif 10 <= change < 30:
            fill = orange_fill
        elif change < 0:
            fill = green_fill
        else:
            fill = None
        if fill:
            for cell in ws[ws.max_row]:
                cell.fill = fill

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name='discrepancies.xlsx')

@app.route('/history')
def history():
    seller_url = request.args.get('seller_url', '')
    if not seller_url:
        return '<p>Не указан продавец.</p>'
    cutoff = (datetime.date.today() - datetime.timedelta(days=7)).isoformat()
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT code, date, price
        FROM price_history
        WHERE code IN (SELECT code FROM sellers WHERE seller_url = ?)
          AND date >= ?
        ORDER BY code, date
    """, (seller_url, cutoff))
    rows = cur.fetchall()
    conn.close()
    if not rows:
        return '<p>Нет данных за неделю по данному продавцу.</p>'

    df = pd.DataFrame(rows, columns=['Код товара', 'Дата', 'Цена'])
    pivot = df.pivot(index='Код товара', columns='Дата', values='Цена').reset_index()
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pivot.to_excel(writer, index=False, sheet_name='История цен')
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=f'history_{seller_url.split("/")[-1]}.xlsx')

@app.route('/download_history/<filename>')
def download_history_file(filename):
    filepath = os.path.join(DOWNLOAD_DIR, filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True, download_name=filename)
    else:
        return '<p>Файл не найден.</p>', 404

if __name__ == '__main__':
    init_db()
    schedule.every(6).hours.do(scheduled_parse)
    scheduler_thread = threading.Thread(target=run_schedule, daemon=True)
    scheduler_thread.start()
    # Слушаем на всех интерфейсах, многопоточный режим
    app.run(host='0.0.0.0', port=5000, threaded=True, debug=False)
