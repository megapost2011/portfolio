import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import sqlite3
import requests
from bs4 import BeautifulSoup
import os
from datetime import datetime
import re
import urllib.parse
import time
import webbrowser
import threading

# バーコードスキャン用ライブラリ
try:
    import cv2
    from pyzbar import pyzbar
    from PIL import Image, ImageTk
    BARCODE_AVAILABLE = True
except ImportError:
    BARCODE_AVAILABLE = False
    print("[WARNING] opencv-python または pyzbar がインストールされていません")
    print("バーコードスキャン機能を使うには以下をインストールしてください:")
    print("pip install opencv-python pyzbar pillow --break-system-packages")

class BookDatabase:
    def __init__(self, db_path="book_manager.db"):
        self.db_path = db_path
        self.init_database()
        self.migrate_database()
    
    def get_connection(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn
    
    def column_exists(self, table_name, column_name):
        conn = self.get_connection()
        cur = conn.cursor()
        cur.execute(f"PRAGMA table_info({table_name})")
        columns = [row[1] for row in cur.fetchall()]
        conn.close()
        return column_name in columns
    
    def migrate_database(self):
        conn = self.get_connection()
        cur = conn.cursor()
        
        new_columns = [
            ('author', 'TEXT'),
            ('publisher', 'TEXT'),
            ('bookoff_price', 'INTEGER'),
            ('bookoff_url', 'TEXT'),
            ('academybook_price', 'INTEGER'),
            ('academybook_url', 'TEXT'),
            ('valuebooks_price', 'INTEGER'),
            ('valuebooks_url', 'TEXT'),
            ('max_buyback_price', 'INTEGER'),
            ('max_buyback_site', 'TEXT'),
            ('purchase_price', 'INTEGER'),
        ]
        
        for column_name, column_type in new_columns:
            if not self.column_exists('books', column_name):
                try:
                    cur.execute(f"ALTER TABLE books ADD COLUMN {column_name} {column_type}")
                    print(f"カラム追加: {column_name}")
                except sqlite3.OperationalError as e:
                    print(f"カラム追加スキップ: {column_name} - {e}")
        
        conn.commit()
        conn.close()
        print("データベースマイグレーション完了")
    
    def init_database(self):
        conn = self.get_connection()
        cur = conn.cursor()
        
        cur.execute("""
            CREATE TABLE IF NOT EXISTS books (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                isbn TEXT UNIQUE,
                title TEXT,
                author TEXT,
                publisher TEXT,
                image_path TEXT,
                amazon_rating REAL,
                amazon_review_count INTEGER,
                amazon_url TEXT,
                bookoff_price INTEGER,
                bookoff_url TEXT,
                academybook_price INTEGER,
                academybook_url TEXT,
                valuebooks_price INTEGER,
                valuebooks_url TEXT,
                max_buyback_price INTEGER,
                max_buyback_site TEXT,
                status TEXT DEFAULT 'owned',
                purchase_date TEXT,
                purchase_price INTEGER,
                sell_date TEXT,
                sell_price INTEGER,
                notes TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        cur.execute("""
            CREATE TABLE IF NOT EXISTS search_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                isbn TEXT,
                title TEXT,
                searched_at TEXT DEFAULT CURRENT_TIMESTAMP,
                found_in_db INTEGER
            )
        """)
        
        cur.execute("CREATE INDEX IF NOT EXISTS idx_isbn ON books(isbn)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_status ON books(status)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_title ON books(title)")
        
        conn.commit()
        conn.close()
    
    def add_book(self, book_data):
        conn = self.get_connection()
        cur = conn.cursor()
        
        try:
            cur.execute("""
                INSERT INTO books (
                    isbn, title, author, publisher, image_path, 
                    amazon_rating, amazon_review_count, amazon_url,
                    bookoff_price, bookoff_url,
                    academybook_price, academybook_url,
                    valuebooks_price, valuebooks_url,
                    max_buyback_price, max_buyback_site,
                    status, purchase_date, purchase_price
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                book_data.get('isbn'),
                book_data.get('title'),
                book_data.get('author'),
                book_data.get('publisher'),
                book_data.get('image_path'),
                book_data.get('amazon_rating'),
                book_data.get('amazon_review_count'),
                book_data.get('amazon_url'),
                book_data.get('bookoff_price'),
                book_data.get('bookoff_url'),
                book_data.get('academybook_price'),
                book_data.get('academybook_url'),
                book_data.get('valuebooks_price'),
                book_data.get('valuebooks_url'),
                book_data.get('max_buyback_price'),
                book_data.get('max_buyback_site'),
                'owned',
                datetime.now().strftime('%Y-%m-%d'),
                book_data.get('purchase_price', 0)
            ))
            
            book_id = cur.lastrowid
            conn.commit()
            return book_id
        except sqlite3.IntegrityError:
            return None
        finally:
            conn.close()
    
    def get_books(self, status=None):
        conn = self.get_connection()
        cur = conn.cursor()
        
        if status:
            cur.execute("SELECT * FROM books WHERE status = ? ORDER BY created_at DESC", (status,))
        else:
            cur.execute("SELECT * FROM books ORDER BY created_at DESC")
        
        books = [dict(row) for row in cur.fetchall()]
        conn.close()
        return books
    
    def update_book(self, book_id, updates):
        conn = self.get_connection()
        cur = conn.cursor()
        
        updates['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        set_clause = ", ".join([f"{k} = ?" for k in updates.keys()])
        values = list(updates.values()) + [book_id]
        
        cur.execute(f"UPDATE books SET {set_clause} WHERE id = ?", values)
        conn.commit()
        conn.close()
    
    def check_duplicate(self, isbn=None, title=None):
        conn = self.get_connection()
        cur = conn.cursor()
        
        results = []
        
        if isbn:
            cur.execute("SELECT * FROM books WHERE isbn = ?", (isbn,))
            results.extend([dict(row) for row in cur.fetchall()])
        
        if title and not results:
            cur.execute("SELECT * FROM books WHERE title LIKE ?", (f"%{title}%",))
            results.extend([dict(row) for row in cur.fetchall()])
        
        cur.execute("""
            INSERT INTO search_history (isbn, title, found_in_db)
            VALUES (?, ?, ?)
        """, (isbn, title, len(results) > 0))
        
        conn.commit()
        conn.close()
        return results
    
    def get_low_rated_books(self, threshold=3.5):
        conn = self.get_connection()
        cur = conn.cursor()
        
        cur.execute("""
            SELECT * FROM books 
            WHERE amazon_rating IS NOT NULL 
            AND amazon_rating < ? 
            AND status = 'owned'
            ORDER BY amazon_rating ASC
        """, (threshold,))
        
        books = [dict(row) for row in cur.fetchall()]
        conn.close()
        return books
    
    def get_stats(self):
        conn = self.get_connection()
        cur = conn.cursor()
        
        cur.execute("""
            SELECT 
                status,
                COUNT(*) as count,
                AVG(CASE WHEN amazon_rating IS NOT NULL THEN amazon_rating END) as avg_rating,
                AVG(CASE WHEN max_buyback_price IS NOT NULL THEN max_buyback_price END) as avg_buyback,
                SUM(CASE WHEN sell_price IS NOT NULL THEN sell_price ELSE 0 END) as total_sales,
                SUM(CASE WHEN purchase_price IS NOT NULL THEN purchase_price ELSE 0 END) as total_purchase
            FROM books
            GROUP BY status
        """)
        
        stats = [dict(row) for row in cur.fetchall()]
        conn.close()
        return stats

class AmazonScraper:
    @staticmethod
    def isbn13_to_isbn10(isbn13):
        isbn13 = str(isbn13).strip()
        if len(isbn13) != 13 or not isbn13.startswith('978'):
            return isbn13
        
        isbn10_base = isbn13[3:-1]
        check = sum((i + 1) * int(x) for i, x in enumerate(isbn10_base)) % 11
        check_digit = 'X' if check == 10 else str(check)
        return isbn10_base + check_digit
    
    @staticmethod
    def search_by_isbn(isbn):
        if not isbn:
            return None
        
        isbn10 = AmazonScraper.isbn13_to_isbn10(isbn)
        url = f"https://www.amazon.co.jp/dp/{isbn10}"
        
        try:
            print(f"[DEBUG] Amazon検索URL: {url}")
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Linux; Android 11; Pixel 5) AppleWebKit/537.36',
                'Accept-Language': 'ja-JP,ja;q=0.9'
            }
            
            response = requests.get(url, headers=headers, timeout=20)
            print(f"[DEBUG] Amazon応答コード: {response.status_code}")
            
            if response.status_code != 200:
                return {'url': url, 'title': None, 'author': None, 'publisher': None, 'rating': None, 'review_count': 0}
            
            soup = BeautifulSoup(response.content, 'html.parser')
            
            title = None
            title_selectors = [
                '#productTitle',
                'span#productTitle',
                'h1#title',
                'h1.product-title',
                '[id="productTitle"]'
            ]
            
            for selector in title_selectors:
                title_elem = soup.select_one(selector)
                if title_elem:
                    title = title_elem.get_text().strip()
                    print(f"[DEBUG] タイトル取得成功: {title[:50]}...")
                    break
            
            if not title:
                print("[DEBUG] タイトル要素が見つかりません")
                all_h1 = soup.find_all('h1')
                print(f"[DEBUG] ページ内のh1要素数: {len(all_h1)}")
                if all_h1:
                    title = all_h1[0].get_text().strip()
                    print(f"[DEBUG] h1から取得: {title[:50]}...")
            
            author = None
            author_selectors = [
                '.author a',
                '.contributorNameID',
                'a.author',
                '[class*="author"] a',
                '.a-link-normal.contributorNameID'
            ]
            
            for selector in author_selectors:
                author_elem = soup.select_one(selector)
                if author_elem:
                    author = author_elem.get_text().strip()
                    print(f"[DEBUG] 著者取得: {author}")
                    break
            
            publisher = None
            details = soup.select('.detail-bullet-list li, #detailBullets_feature_div li, #detailBulletsWrapper_feature_div li')
            for detail in details:
                text = detail.get_text()
                if '出版社' in text:
                    try:
                        publisher = text.split('出版社')[1].split('(')[0].strip().replace(':', '').strip()
                        print(f"[DEBUG] 出版社取得: {publisher}")
                    except:
                        pass
                    break
            
            rating = None
            rating_selectors = [
                'span.a-icon-alt',
                'i.a-icon-star span.a-icon-alt',
                '[data-hook="rating-out-of-text"]',
                '.a-star-5'
            ]
            
            for selector in rating_selectors:
                rating_elem = soup.select_one(selector)
                if rating_elem:
                    rating_text = rating_elem.get_text()
                    match = re.search(r'(\d+\.?\d*)', rating_text)
                    if match:
                        rating = float(match.group(1))
                        print(f"[DEBUG] 評価取得: {rating}")
                        break
            
            review_count = 0
            review_selectors = [
                '#acrCustomerReviewText',
                '[data-hook="total-review-count"]',
                'span#acrCustomerReviewText'
            ]
            
            for selector in review_selectors:
                review_elem = soup.select_one(selector)
                if review_elem:
                    review_text = review_elem.get_text()
                    numbers = re.findall(r'\d+', review_text.replace(',', ''))
                    if numbers:
                        review_count = int(numbers[0])
                        print(f"[DEBUG] レビュー数取得: {review_count}")
                        break
            
            if review_count == 0:
                review_links = soup.select('a[href*="#customerReviews"]')
                for link in review_links:
                    text = link.get_text()
                    numbers = re.findall(r'\d+', text.replace(',', ''))
                    if numbers:
                        review_count = int(numbers[0])
                        break
            
            return {
                'title': title,
                'author': author,
                'publisher': publisher,
                'rating': rating,
                'review_count': review_count,
                'url': url
            }
        
        except Exception as e:
            print(f"[ERROR] Amazon検索エラー: {e}")
            import traceback
            traceback.print_exc()
            return {'url': url, 'title': None, 'author': None, 'publisher': None, 'rating': None, 'review_count': 0}
    
    @staticmethod
    def search_by_title(title):
        if not title:
            return None
        
        try:
            query = urllib.parse.quote(title + ' 本')
            url = f"https://www.amazon.co.jp/s?k={query}&i=stripbooks"
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Linux; Android 11) AppleWebKit/537.36',
                'Accept-Language': 'ja-JP,ja;q=0.9'
            }
            
            response = requests.get(url, headers=headers, timeout=20)
            
            if response.status_code != 200:
                return None
            
            soup = BeautifulSoup(response.content, 'html.parser')
            first_result = soup.select_one('[data-component-type="s-search-result"]')
            
            if not first_result:
                return None
            
            title_elem = first_result.select_one('h2 a span')
            result_title = title_elem.get_text().strip() if title_elem else title
            
            rating = None
            rating_elem = first_result.select_one('i.a-icon-star-small span.a-icon-alt')
            if rating_elem:
                rating_text = rating_elem.get_text()
                match = re.search(r'(\d+\.?\d*)', rating_text)
                if match:
                    rating = float(match.group(1))
            
            review_count = 0
            review_elem = first_result.select_one('span[aria-label*="個の評価"]')
            if review_elem:
                aria_label = review_elem.get('aria-label', '')
                numbers = re.findall(r'\d+', aria_label.replace(',', ''))
                if numbers:
                    review_count = int(numbers[0])
            
            link_elem = first_result.select_one('h2 a')
            product_url = None
            if link_elem:
                href = link_elem.get('href')
                product_url = f"https://www.amazon.co.jp{href}"
            
            return {
                'title': result_title,
                'author': None,
                'publisher': None,
                'rating': rating,
                'review_count': review_count,
                'url': product_url
            }
        
        except Exception as e:
            print(f"Amazon書籍名検索エラー: {e}")
            return None

class BuybackURLManager:
    @staticmethod
    def get_buyback_urls(isbn):
        """各買取サイトのURL（手動確認用）を生成"""
        results = {
            'bookoff': None,
            'academybook': None,
            'valuebooks': None,
        }
        
        if not isbn:
            return results
        
        results['bookoff'] = "https://www.bookoffonline.co.jp/boleccontent/bolbuysearch/buysearch/display"
        results['academybook'] = "https://www.academybook.net/smt/item/item05.html?a8=OZ2qUZfm4k9qHKe3-B3hPktk0lL.tk3kXB3qDY3UQ_ov9k2PqkMOWktO-lMZPkMjoiD3-lMjal2QSG2mX_ol9N2OoZSe4G21qGc.xs00000014429002"
        results['valuebooks'] = "https://www.valuebooks.jp/endpaper/3530/"
        
        return results

class BarcodeScannerWindow:
    """バーコードスキャナーウィンドウ"""
    def __init__(self, parent, callback):
        self.parent = parent
        self.callback = callback
        self.window = tk.Toplevel(parent)
        self.window.title("📷 ISBNバーコードスキャン")
        self.window.geometry("640x520")
        
        self.scanning = True
        self.camera = None
        self.last_scan_time = 0
        self.scan_cooldown = 2.0  # 2秒間のクールダウン
        
        # カメラプレビュー
        self.camera_label = ttk.Label(self.window)
        self.camera_label.pack(pady=10, padx=10, fill='both', expand=True)
        
        # ステータス表示
        self.status_var = tk.StringVar(value="カメラ起動中...")
        ttk.Label(self.window, textvariable=self.status_var, 
                 font=('', 12), foreground='blue').pack(pady=5)
        
        # ISBNバーコードを本の背表紙に向ける
        guide_text = "📖 ISBNバーコードをカメラに映してください\n（本の裏表紙または背表紙にあります）"
        ttk.Label(self.window, text=guide_text, foreground='gray').pack(pady=5)
        
        # 閉じるボタン
        ttk.Button(self.window, text="✖️ 閉じる", 
                  command=self.close).pack(pady=10, fill='x', padx=20)
        
        self.window.protocol("WM_DELETE_WINDOW", self.close)
        
        # カメラ起動
        self.start_camera()
    
    def start_camera(self):
        """カメラを起動"""
        try:
            # カメラを開く（0は通常背面カメラ、1はフロントカメラ）
            self.camera = cv2.VideoCapture(0)
            
            if not self.camera.isOpened():
                # カメラ0が開けない場合、カメラ1を試す
                self.camera = cv2.VideoCapture(1)
            
            if not self.camera.isOpened():
                raise Exception("カメラを開けませんでした")
            
            # カメラ解像度設定
            self.camera.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
            self.camera.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
            
            self.status_var.set("バーコードをスキャン中...")
            
            # フレーム更新を開始
            self.update_frame()
            
        except Exception as e:
            self.status_var.set(f"カメラエラー: {e}")
            messagebox.showerror("エラー", 
                f"カメラを起動できませんでした\n\n{e}\n\n"
                "カメラへのアクセス権限を確認してください")
    
    def update_frame(self):
        """カメラフレームを更新してバーコードをスキャン"""
        if not self.scanning or not self.camera:
            return
        
        try:
            ret, frame = self.camera.read()
            
            if ret:
                # バーコードをデコード
                barcodes = pyzbar.decode(frame)
                
                current_time = time.time()
                
                for barcode in barcodes:
                    # バーコードデータを取得
                    barcode_data = barcode.data.decode('utf-8')
                    barcode_type = barcode.type
                    
                    print(f"[SCAN] Type: {barcode_type}, Data: {barcode_data}")
                    
                    # ISBNバーコード（EAN-13）を検出
                    if barcode_type == 'EAN13' and len(barcode_data) == 13:
                        # クールダウンチェック
                        if current_time - self.last_scan_time > self.scan_cooldown:
                            self.last_scan_time = current_time
                            
                            # バーコードを四角で囲む
                            (x, y, w, h) = barcode.rect
                            cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 3)
                            
                            # ISBNテキストを表示
                            text = f"ISBN: {barcode_data}"
                            cv2.putText(frame, text, (x, y - 10), 
                                      cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 0), 2)
                            
                            self.status_var.set(f"✅ ISBN検出: {barcode_data}")
                            
                            # コールバック実行
                            self.window.after(100, lambda: self.on_isbn_detected(barcode_data))
                            break
                    
                    # バーコードを検出したが、ISBNではない場合
                    else:
                        (x, y, w, h) = barcode.rect
                        cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 2)
                        cv2.putText(frame, f"{barcode_type}", (x, y - 10), 
                                  cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 0, 255), 2)
                
                # フレームをTkinterで表示可能な形式に変換
                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                frame = cv2.resize(frame, (640, 480))
                
                img = Image.fromarray(frame)
                imgtk = ImageTk.PhotoImage(image=img)
                
                self.camera_label.imgtk = imgtk
                self.camera_label.configure(image=imgtk)
            
            # 次のフレームを予約（約30fps）
            self.window.after(33, self.update_frame)
            
        except Exception as e:
            print(f"[ERROR] Frame update error: {e}")
            self.status_var.set(f"エラー: {e}")
    
    def on_isbn_detected(self, isbn):
        """ISBNが検出されたときの処理"""
        print(f"[INFO] ISBN detected: {isbn}")
        self.callback(isbn)
        self.close()
    
    def close(self):
        """ウィンドウを閉じる"""
        self.scanning = False
        
        if self.camera:
            self.camera.release()
        
        self.window.destroy()

class JapaneseInputDialog:
    @staticmethod
    def get_input(parent, title, prompt, initial_value=""):
        dialog = tk.Toplevel(parent)
        dialog.title(title)
        dialog.geometry("350x200")
        dialog.transient(parent)
        dialog.grab_set()
        
        result = {'value': None}
        
        ttk.Label(dialog, text=prompt, wraplength=300).pack(pady=10, padx=10)
        
        text_frame = ttk.Frame(dialog)
        text_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        text_widget = tk.Text(text_frame, height=4, width=40, font=('', 12))
        text_widget.pack(fill='both', expand=True)
        
        if initial_value:
            text_widget.insert('1.0', initial_value)
        
        text_widget.focus_set()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill='x', padx=10, pady=10)
        
        def on_ok():
            result['value'] = text_widget.get('1.0', 'end-1c').strip()
            dialog.destroy()
        
        def on_cancel():
            result['value'] = None
            dialog.destroy()
        
        ttk.Button(btn_frame, text="OK", command=on_ok).pack(side='left', expand=True, fill='x', padx=2)
        ttk.Button(btn_frame, text="キャンセル", command=on_cancel).pack(side='left', expand=True, fill='x', padx=2)
        
        help_text = "※ Gboard/Simejiで日本語入力"
        ttk.Label(dialog, text=help_text, foreground='gray', font=('', 8)).pack(pady=5)
        
        parent.wait_window(dialog)
        
        return result['value']

class BookManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📚 蔵書管理Pro")
        self.root.geometry("420x750")
        
        self.db = BookDatabase()
        
        self.image_dir = "book_images"
        os.makedirs(self.image_dir, exist_ok=True)
        
        self.current_books = []
        self.low_rated_books = []
        self.current_amazon_url = None
        self.current_buyback_urls = {}
        self.current_isbn = None
        
        self.create_widgets()
        self.show_welcome()
    
    def show_welcome(self):
        stats = self.db.get_stats()
        total = sum(s['count'] for s in stats)
        
        welcome_msg = (
            f"📚 蔵書管理システム Pro\n\n"
            f"✅ Amazon評価＋レビュー数\n"
            f"✅ ISBNバーコードスキャン\n"
            f"✅ 買取サイトへのリンク提供\n"
            f"  ・ブックオフオンライン\n"
            f"  ・専門書アカデミー\n"
            f"  ・バリューブックス\n\n"
            f"登録済み: {total}冊\n\n"
        )
        
        if not BARCODE_AVAILABLE:
            welcome_msg += "⚠️ バーコードスキャン機能が無効です\n以下をインストールしてください:\n"
            welcome_msg += "pip install opencv-python pyzbar pillow --break-system-packages"
        else:
            welcome_msg += "💡 バーコードボタンで簡単登録！"
        
        messagebox.showinfo("起動", welcome_msg)
    
    def create_widgets(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.create_register_tab()
        self.create_list_tab()
        self.create_sell_tab()
        self.create_check_tab()
        self.create_stats_tab()
    
    def create_register_tab(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text='📝 登録')
        
        # バーコードスキャンボタン（目立つように配置）
        if BARCODE_AVAILABLE:
            barcode_btn = ttk.Button(tab, text="📷 ISBNバーコードスキャン", 
                                     command=self.open_barcode_scanner)
            barcode_btn.pack(fill='x', pady=10)
        else:
            barcode_frame = ttk.Frame(tab)
            barcode_frame.pack(fill='x', pady=10)
            
            ttk.Label(barcode_frame, text="⚠️ バーコードスキャン機能無効", 
                     foreground='red', font=('', 9)).pack()
            ttk.Label(barcode_frame, text="opencv-python, pyzbarをインストール", 
                     foreground='gray', font=('', 8)).pack()
        
        isbn_frame = ttk.Frame(tab)
        isbn_frame.pack(fill='x', pady=5)
        
        ttk.Label(isbn_frame, text="ISBN:").pack(side='left')
        self.isbn_var = tk.StringVar()
        isbn_entry = ttk.Entry(isbn_frame, textvariable=self.isbn_var)
        isbn_entry.pack(side='left', fill='x', expand=True, padx=5)
        
        ttk.Button(isbn_frame, text="🔍", width=3,
                   command=lambda: self.search_by_isbn()).pack(side='left')
        
        title_frame = ttk.Frame(tab)
        title_frame.pack(fill='x', pady=5)
        
        ttk.Label(title_frame, text="タイトル:").pack(side='left')
        self.title_var = tk.StringVar()
        title_entry = ttk.Entry(title_frame, textvariable=self.title_var)
        title_entry.pack(side='left', fill='x', expand=True, padx=5)
        
        ttk.Button(title_frame, text="✏️", width=3,
                   command=self.input_title).pack(side='left')
        ttk.Button(title_frame, text="🔍", width=3,
                   command=self.search_by_title).pack(side='left')
        
        self.author_var = tk.StringVar()
        self.publisher_var = tk.StringVar()
        
        info_frame = ttk.Frame(tab)
        info_frame.pack(fill='x', pady=5)
        
        ttk.Label(info_frame, text="著者:").grid(row=0, column=0, sticky='w')
        ttk.Label(info_frame, textvariable=self.author_var, foreground='blue').grid(row=0, column=1, sticky='w')
        
        ttk.Label(info_frame, text="出版:").grid(row=1, column=0, sticky='w')
        ttk.Label(info_frame, textvariable=self.publisher_var, foreground='blue').grid(row=1, column=1, sticky='w')
        
        amazon_frame = ttk.LabelFrame(tab, text="Amazon情報", padding=5)
        amazon_frame.pack(fill='x', pady=5)
        
        self.amazon_rating_var = tk.StringVar(value="未取得")
        self.amazon_reviews_var = tk.StringVar(value="0件")
        
        ttk.Label(amazon_frame, text="評価:").grid(row=0, column=0, sticky='w')
        ttk.Label(amazon_frame, textvariable=self.amazon_rating_var,
                 foreground='red', font=('', 11, 'bold')).grid(row=0, column=1, sticky='w', padx=5)
        
        ttk.Label(amazon_frame, text="レビュー:").grid(row=1, column=0, sticky='w')
        ttk.Label(amazon_frame, textvariable=self.amazon_reviews_var,
                 foreground='green').grid(row=1, column=1, sticky='w', padx=5)
        
        buyback_frame = ttk.LabelFrame(tab, text="買取相場確認", padding=5)
        buyback_frame.pack(fill='x', pady=5)
        
        ttk.Label(buyback_frame, text="※サイトで手動確認してください", 
                 foreground='gray', font=('', 8)).pack(pady=2)
        
        ttk.Button(buyback_frame, text="🔗 ブックオフで確認",
                  command=lambda: self.open_buyback_url('bookoff')).pack(fill='x', pady=2)
        ttk.Button(buyback_frame, text="🔗 専門書アカデミーで確認",
                  command=lambda: self.open_buyback_url('academybook')).pack(fill='x', pady=2)
        ttk.Button(buyback_frame, text="🔗 バリューブックスで確認",
                  command=lambda: self.open_buyback_url('valuebooks')).pack(fill='x', pady=2)
        
        ttk.Button(buyback_frame, text="📋 ISBNをコピー",
                  command=self.copy_isbn).pack(fill='x', pady=2)
        
        price_frame = ttk.Frame(tab)
        price_frame.pack(fill='x', pady=5)
        
        ttk.Label(price_frame, text="買取価格(手入力):").pack(side='left')
        self.manual_price_var = tk.StringVar(value="0")
        ttk.Entry(price_frame, textvariable=self.manual_price_var, width=10).pack(side='left', padx=5)
        ttk.Label(price_frame, text="円").pack(side='left')
        
        ttk.Label(tab, text="メモ:").pack(anchor='w')
        self.notes_text = scrolledtext.ScrolledText(tab, height=2)
        self.notes_text.pack(fill='both', expand=True, pady=2)
        
        ttk.Button(tab, text="✅ 登録する",
                   command=self.register_book).pack(pady=5, fill='x')
        
        ttk.Button(tab, text="🗑️ クリア",
                   command=self.clear_form).pack(fill='x')
        
        hint = "💡 バーコードスキャンで簡単登録"
        ttk.Label(tab, text=hint, foreground='gray', font=('', 8)).pack(pady=5)
    
    def open_barcode_scanner(self):
        """バーコードスキャナーを開く"""
        if not BARCODE_AVAILABLE:
            messagebox.showerror("エラー", 
                "バーコードスキャン機能が利用できません\n\n"
                "以下のコマンドでライブラリをインストールしてください:\n"
                "pip install opencv-python pyzbar pillow --break-system-packages")
            return
        
        def on_isbn_scanned(isbn):
            """ISBNがスキャンされたときの処理"""
            self.isbn_var.set(isbn)
            self.current_isbn = isbn
            
            # 自動的にAmazon検索を実行
            self.root.after(100, self.search_by_isbn)
        
        # バーコードスキャナーウィンドウを開く
        BarcodeScannerWindow(self.root, on_isbn_scanned)
    
    def copy_isbn(self):
        """ISBNをクリップボードにコピー"""
        isbn = self.current_isbn or self.isbn_var.get().strip()
        if isbn:
            self.root.clipboard_clear()
            self.root.clipboard_append(isbn)
            messagebox.showinfo("コピー完了", f"ISBN: {isbn}\nをクリップボードにコピーしました")
        else:
            messagebox.showwarning("警告", "ISBNを検索してください")
    
    def open_buyback_url(self, site):
        """買取サイトをブラウザで開く"""
        if site in self.current_buyback_urls and self.current_buyback_urls[site]:
            try:
                webbrowser.open(self.current_buyback_urls[site])
                
                if site == 'bookoff':
                    messagebox.showinfo("ブックオフ 使い方", 
                        "ブックオフサイトが開きます\n\n"
                        "1. ISBN入力欄にISBNを貼り付け\n"
                        "2. 検索ボタンをタップ\n"
                        "3. 買取価格を確認\n\n"
                        "※「📋 ISBNをコピー」ボタンで\n"
                        "ISBNをコピーできます")
                elif site == 'academybook':
                    messagebox.showinfo("専門書アカデミー 使い方", 
                        "専門書アカデミーサイトが開きます\n\n"
                        "1. ISBN入力欄にISBNを貼り付け\n"
                        "2. 検索ボタンをタップ\n"
                        "3. 買取価格を確認\n\n"
                        "※専門書・教科書・参考書などの\n"
                        "買取価格が高めのサイトです")
                elif site == 'valuebooks':
                    messagebox.showinfo("バリューブックス 使い方", 
                        "バリューブックスサイトが開きます\n\n"
                        "1. ISBN入力欄にISBNを貼り付け\n"
                        "2. 検索ボタンをタップ\n"
                        "3. 買取価格を確認\n\n"
                        "※古書・絶版本・希少本などの\n"
                        "買取に強いサイトです")
            except:
                messagebox.showinfo("URL", 
                    f"以下のURLをブラウザで開いてください:\n\n{self.current_buyback_urls[site]}")
        else:
            messagebox.showwarning("警告", "ISBNを検索してからご利用ください")
    
    def input_title(self):
        current = self.title_var.get()
        result = JapaneseInputDialog.get_input(
            self.root,
            "タイトル入力",
            "書籍のタイトルを入力してください:",
            current
        )
        
        if result:
            self.title_var.set(result)
    
    def search_by_isbn(self):
        isbn = self.isbn_var.get().strip()
        
        if not isbn:
            messagebox.showwarning("警告", "ISBNを入力してください")
            return
        
        self.current_isbn = isbn
        
        messagebox.showinfo("検索中", "Amazon情報取得中...")
        self.root.update()
        
        amazon_data = AmazonScraper.search_by_isbn(isbn)
        
        if amazon_data and amazon_data.get('url'):
            self.current_amazon_url = amazon_data.get('url')
            self.title_var.set(amazon_data.get('title') or '')
            self.author_var.set(amazon_data.get('author') or '')
            self.publisher_var.set(amazon_data.get('publisher') or '')
            
            rating = amazon_data.get('rating')
            reviews = amazon_data.get('review_count', 0)
            
            self.amazon_rating_var.set(f"★{rating}" if rating else "未取得")
            self.amazon_reviews_var.set(f"{reviews}件")
        
        self.current_buyback_urls = BuybackURLManager.get_buyback_urls(isbn)
        
        if amazon_data and amazon_data.get('title'):
            rating = amazon_data.get('rating')
            reviews = amazon_data.get('review_count', 0)
            messagebox.showinfo("完了",
                f"✅ 情報取得完了\n\n"
                f"{amazon_data.get('title', '')[:30]}\n"
                f"評価: ★{rating if rating else '未取得'} ({reviews}件)\n\n"
                f"買取サイトボタンで価格確認できます")
        else:
            messagebox.showwarning("警告",
                f"Amazon情報が取得できませんでした\n"
                f"URLは保存されています\n"
                f"タイトルを手入力してください")
    
    def search_by_title(self):
        title = self.title_var.get().strip()
        
        if not isbn:
            messagebox.showwarning("警告", "タイトルを入力してください")
            return
        
        messagebox.showinfo("検索中", f"「{title}」を検索中...")
        self.root.update()
        
        amazon_data = AmazonScraper.search_by_title(title)
        
        if amazon_data:
            self.title_var.set(amazon_data.get('title', title))
            self.current_amazon_url = amazon_data.get('url')
            
            rating = amazon_data.get('rating')
            reviews = amazon_data.get('review_count', 0)
            
            self.amazon_rating_var.set(f"★{rating}" if rating else "未取得")
            self.amazon_reviews_var.set(f"{reviews}件")
            
            messagebox.showinfo("成功", f"✅ 検索完了\n\n評価: ★{rating if rating else '未取得'} ({reviews}件)")
        else:
            messagebox.showerror("エラー", "書籍が見つかりませんでした")
    
    def register_book(self):
        isbn = self.isbn_var.get().strip()
        title = self.title_var.get().strip()
        author = self.author_var.get().strip()
        publisher = self.publisher_var.get().strip()
        notes = self.notes_text.get('1.0', 'end').strip()
        
        if not isbn and not title:
            messagebox.showwarning("警告", "ISBNまたはタイトルを入力してください")
            return
        
        if isbn:
            duplicates = self.db.check_duplicate(isbn=isbn)
            if duplicates:
                if not messagebox.askyesno("重複",
                    f"ISBN: {isbn}\nは既に登録されています。\n続行しますか？"):
                    return
        
        rating = None
        rating_text = self.amazon_rating_var.get()
        if '★' in rating_text:
            try:
                rating = float(rating_text.replace('★', ''))
            except:
                pass
        
        review_count = 0
        review_text = self.amazon_reviews_var.get()
        numbers = re.findall(r'\d+', review_text.replace(',', ''))
        if numbers:
            review_count = int(numbers[0])
        
        manual_price = None
        try:
            manual_price = int(self.manual_price_var.get())
        except:
            pass
        
        book_data = {
            'isbn': isbn,
            'title': title,
            'author': author or None,
            'publisher': publisher or None,
            'amazon_rating': rating,
            'amazon_review_count': review_count,
            'amazon_url': self.current_amazon_url,
            'bookoff_price': None,
            'bookoff_url': self.current_buyback_urls.get('bookoff'),
            'academybook_price': None,
            'academybook_url': self.current_buyback_urls.get('academybook'),
            'valuebooks_price': None,
            'valuebooks_url': self.current_buyback_urls.get('valuebooks'),
            'max_buyback_price': manual_price,
            'max_buyback_site': '手入力' if manual_price else None,
            'image_path': None,
            'purchase_price': 0
        }
        
        book_id = self.db.add_book(book_data)
        
        if book_id:
            messagebox.showinfo("成功",
                f"✅ 登録完了！\n\n"
                f"{title}\n"
                f"評価: {rating_text} ({review_count}件)\n"
                f"買取価格: {manual_price or 0}円")
            self.clear_form()
        else:
            messagebox.showerror("エラー", "登録に失敗しました（重複の可能性）")
    
    def create_list_tab(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text='📋 一覧')
        
        filter_frame = ttk.Frame(tab)
        filter_frame.pack(fill='x', pady=5)
        
        self.filter_var = tk.StringVar(value='all')
        ttk.Radiobutton(filter_frame, text="全て", variable=self.filter_var,
                       value='all', command=self.load_books).pack(side='left', padx=2)
        ttk.Radiobutton(filter_frame, text="所有", variable=self.filter_var,
                       value='owned', command=self.load_books).pack(side='left', padx=2)
        ttk.Radiobutton(filter_frame, text="売却済", variable=self.filter_var,
                       value='sold', command=self.load_books).pack(side='left', padx=2)
        
        list_frame = ttk.Frame(tab)
        list_frame.pack(fill='both', expand=True, pady=5)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side='right', fill='y')
        
        self.book_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        self.book_listbox.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=self.book_listbox.yview)
        
        self.book_listbox.bind('<<ListboxSelect>>', self.on_book_select)
        
        ttk.Label(tab, text="詳細:").pack(anchor='w')
        self.detail_text = scrolledtext.ScrolledText(tab, height=10)
        self.detail_text.pack(fill='both', expand=True, pady=5)
        
        btn_frame = ttk.Frame(tab)
        btn_frame.pack(fill='x')
        
        ttk.Button(btn_frame, text="🔄 更新",
                   command=self.load_books).pack(side='left', expand=True, fill='x', padx=2)
        ttk.Button(btn_frame, text="🗑️ 削除",
                   command=self.delete_book).pack(side='left', expand=True, fill='x', padx=2)
    
    def create_sell_tab(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text='💰 売却')
        
        ttk.Label(tab, text="売却候補", font=('', 11, 'bold')).pack(pady=5)
        
        criteria_frame = ttk.Frame(tab)
        criteria_frame.pack(fill='x', pady=5)
        
        ttk.Label(criteria_frame, text="評価:").pack(side='left')
        self.rating_threshold = tk.DoubleVar(value=3.5)
        ttk.Spinbox(criteria_frame, from_=1.0, to=5.0, increment=0.5,
                   textvariable=self.rating_threshold, width=5).pack(side='left', padx=2)
        ttk.Label(criteria_frame, text="以下").pack(side='left')
        
        ttk.Button(criteria_frame, text="🔍 検索",
                  command=self.find_low_rated).pack(side='left', padx=5, expand=True, fill='x')
        
        list_frame = ttk.Frame(tab)
        list_frame.pack(fill='both', expand=True, pady=5)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side='right', fill='y')
        
        self.sell_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        self.sell_listbox.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=self.sell_listbox.yview)
        
        price_frame = ttk.Frame(tab)
        price_frame.pack(fill='x', pady=5)
        
        ttk.Label(price_frame, text="売却価格:").pack(side='left')
        self.sell_price_entry = ttk.Entry(price_frame, width=10)
        self.sell_price_entry.pack(side='left', padx=2)
        self.sell_price_entry.insert(0, "0")
        ttk.Label(price_frame, text="円").pack(side='left')
        
        ttk.Button(tab, text="✅ 売却済みにする",
                  command=self.mark_as_sold).pack(fill='x', pady=5)
    
    def create_check_tab(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text='🔍 重複確認')
        
        ttk.Label(tab, text="購入前チェック", font=('', 11, 'bold')).pack(pady=5)
        
        isbn_frame = ttk.Frame(tab)
        isbn_frame.pack(fill='x', pady=5)
        
        ttk.Label(isbn_frame, text="ISBN:").pack(side='left')
        self.check_isbn_var = tk.StringVar()
        ttk.Entry(isbn_frame, textvariable=self.check_isbn_var).pack(side='left', fill='x', expand=True, padx=5)
        
        title_frame = ttk.Frame(tab)
        title_frame.pack(fill='x', pady=5)
        
        ttk.Label(title_frame, text="タイトル:").pack(side='left')
        self.check_title_var = tk.StringVar()
        ttk.Entry(title_frame, textvariable=self.check_title_var).pack(side='left', fill='x', expand=True, padx=5)
        ttk.Button(title_frame, text="✏️", width=3,
                  command=self.input_check_title).pack(side='left')
        
        ttk.Button(tab, text="🔍 重複チェック",
                  command=self.check_duplicate).pack(pady=10, fill='x')
        
        self.check_result_text = scrolledtext.ScrolledText(tab)
        self.check_result_text.pack(fill='both', expand=True, pady=5)
    
    def input_check_title(self):
        current = self.check_title_var.get()
        result = JapaneseInputDialog.get_input(
            self.root,
            "タイトル入力",
            "チェックする書籍のタイトルを入力:",
            current
        )
        
        if result:
            self.check_title_var.set(result)
    
    def create_stats_tab(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text='📊 統計')
        
        self.stats_text = scrolledtext.ScrolledText(tab)
        self.stats_text.pack(fill='both', expand=True, pady=5)
        
        ttk.Button(tab, text="🔄 更新", command=self.load_stats).pack(fill='x')
        
        self.load_stats()
    
    def load_books(self):
        status = None if self.filter_var.get() == 'all' else self.filter_var.get()
        books = self.db.get_books(status=status)
        
        self.book_listbox.delete(0, 'end')
        self.current_books = books
        
        for book in books:
            rating = f"★{book['amazon_rating']}" if book['amazon_rating'] else "未評価"
            buyback = f"[{book['max_buyback_price']}円]" if book['max_buyback_price'] else ""
            emoji = "📗" if book['status'] == 'owned' else "📕"
            self.book_listbox.insert('end',
                f"{emoji} {book['title'][:18]} {rating} {buyback}")
    
    def on_book_select(self, event):
        sel = self.book_listbox.curselection()
        if not sel:
            return
        
        book = self.current_books[sel[0]]
        
        detail = f"""【{book['title']}】

著者: {book['author'] or '不明'}
出版社: {book['publisher'] or '不明'}
ISBN: {book['isbn'] or '未登録'}

Amazon評価: {f"★{book['amazon_rating']}" if book['amazon_rating'] else '未取得'}
レビュー数: {book['amazon_review_count'] or 0}件
URL: {book['amazon_url'] or '-'}

【買取サイト】
ブックオフ: {book['bookoff_url'] or '-'}
専門書アカデミー: {book['academybook_url'] or '-'}
バリューブックス: {book['valuebooks_url'] or '-'}
買取価格: {f"{book['max_buyback_price']}円 ({book['max_buyback_site']})" if book['max_buyback_price'] else '-'}

ステータス: {book['status']}
購入日: {book['purchase_date'] or '不明'}
登録日: {book['created_at'][:10] if book['created_at'] else '不明'}
"""
        
        self.detail_text.delete('1.0', 'end')
        self.detail_text.insert('1.0', detail)
    
    def delete_book(self):
        sel = self.book_listbox.curselection()
        if not sel:
            messagebox.showwarning("警告", "本を選択してください")
            return
        
        book = self.current_books[sel[0]]
        
        if messagebox.askyesno("確認", f"削除しますか？\n\n{book['title']}"):
            conn = self.db.get_connection()
            cur = conn.cursor()
            cur.execute("DELETE FROM books WHERE id = ?", (book['id'],))
            conn.commit()
            conn.close()
            
            messagebox.showinfo("削除", "削除しました")
            self.load_books()
    
    def find_low_rated(self):
        threshold = self.rating_threshold.get()
        books = self.db.get_low_rated_books(threshold)
        
        self.sell_listbox.delete(0, 'end')
        self.low_rated_books = books
        
        for book in books:
            buyback = f"[{book['max_buyback_price']}円]" if book['max_buyback_price'] else ""
            self.sell_listbox.insert('end',
                f"{book['title'][:22]} ★{book['amazon_rating']} {buyback}")
        
        messagebox.showinfo("結果", f"{len(books)}冊見つかりました")
    
    def mark_as_sold(self):
        sel = self.sell_listbox.curselection()
        if not sel:
            messagebox.showwarning("警告", "本を選択してください")
            return
        
        book = self.low_rated_books[sel[0]]
        
        try:
            price = int(self.sell_price_entry.get())
        except:
            price = 0
        
        if messagebox.askyesno("確認",
            f"売却済みにしますか？\n\n{book['title']}\n売却価格: {price}円"):
            updates = {
                'status': 'sold',
                'sell_date': datetime.now().strftime('%Y-%m-%d'),
                'sell_price': price
            }
            self.db.update_book(book['id'], updates)
            messagebox.showinfo("完了", "売却済みにしました")
            self.find_low_rated()
            self.load_stats()
    
    def check_duplicate(self):
        isbn = self.check_isbn_var.get().strip()
        title = self.check_title_var.get().strip()
        
        if not isbn and not title:
            messagebox.showwarning("警告", "ISBNまたはタイトルを入力")
            return
        
        results = self.db.check_duplicate(isbn, title)
        
        self.check_result_text.delete('1.0', 'end')
        
        if results:
            text = f"⚠️ {len(results)}冊が既に登録されています\n\n"
            for book in results:
                rating = f"★{book['amazon_rating']}" if book['amazon_rating'] else "未評価"
                buyback = f"買取{book['max_buyback_price']}円" if book['max_buyback_price'] else ""
                text += f"・{book['title']}\n  {rating} ({book['amazon_review_count'] or 0}件)\n"
                text += f"  {book['status']} {buyback}\n\n"
            messagebox.showwarning("重複", "既に所有しています")
        else:
            text = "✅ 重複なし\n購入できます！"
            messagebox.showinfo("OK", "重複なし")
        
        self.check_result_text.insert('1.0', text)
    
    def load_stats(self):
        stats = self.db.get_stats()
        books = self.db.get_books()
        
        text = "="*35 + "\n📊 蔵書統計レポート\n" + "="*35 + "\n\n"
        
        total = sum(s['count'] for s in stats)
        text += f"総数: {total}冊\n\n"
        
        for stat in stats:
            name = {"owned": "所有中", "sold": "売却済"}.get(stat['status'], stat['status'])
            avg = f"{stat['avg_rating']:.2f}" if stat['avg_rating'] else "N/A"
            avg_buyback = f"{stat['avg_buyback']:.0f}" if stat['avg_buyback'] else "N/A"
            text += f"{name}: {stat['count']}冊\n"
            text += f"  平均評価: ★{avg}\n"
            text += f"  平均買取: {avg_buyback}円\n"
        
        total_purchase = sum(s.get('total_purchase', 0) or 0 for s in stats)
        total_sales = sum(s.get('total_sales', 0) or 0 for s in stats)
        
        text += f"\n購入総額: {total_purchase:,}円\n"
        text += f"売却総額: {total_sales:,}円\n"
        
        buyback_books = [b for b in books if b['max_buyback_price'] and b['status'] == 'owned']
        if buyback_books:
            top5_buyback = sorted(buyback_books, key=lambda x: x['max_buyback_price'], reverse=True)[:5]
            text += "\n" + "="*35 + "\n【買取価格 TOP5】\n"
            for i, book in enumerate(top5_buyback, 1):
                text += f"{i}. {book['title'][:25]}\n"
                text += f"   {book['max_buyback_price']}円\n"
        
        rated_books = [b for b in books if b['amazon_rating']]
        if rated_books:
            top5 = sorted(rated_books, key=lambda x: x['amazon_rating'], reverse=True)[:5]
            text += "\n" + "="*35 + "\n【高評価 TOP5】\n"
            for i, book in enumerate(top5, 1):
                text += f"{i}. {book['title'][:25]}\n"
                text += f"   ★{book['amazon_rating']} ({book['amazon_review_count']}件)\n"
        
        self.stats_text.delete('1.0', 'end')
        self.stats_text.insert('1.0', text)
    
    def clear_form(self):
        self.isbn_var.set('')
        self.title_var.set('')
        self.author_var.set('')
        self.publisher_var.set('')
        self.amazon_rating_var.set('未取得')
        self.amazon_reviews_var.set('0件')
        self.manual_price_var.set('0')
        self.current_amazon_url = None
        self.current_buyback_urls = {}
        self.current_isbn = None
        self.notes_text.delete('1.0', 'end')

if __name__ == '__main__':
    root = tk.Tk()
    app = BookManagerApp(root)
    root.mainloop()