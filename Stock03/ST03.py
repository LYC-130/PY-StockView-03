import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import yfinance as yf
import os
import json
import webbrowser
from datetime import datetime
import threading
from functools import lru_cache

CONFIG_FILE = "portfolio_config.json"
REFRESH_INTERVAL = 50000  # 50秒刷新一次

class PortfolioTab(ttk.Frame):
    def __init__(self, master, filename, pane_side, main_app):  # 正确定义4个参数
        super().__init__(master)
        # 新增缓存相关属性
        self.cache_valid = False
        self.cached_items = []
        
        self.main_app = main_app  # 新增主應用程式引用
        self.filename = filename
        self.pane_side = pane_side
        self.stocks = []
        
        self.create_widgets()
        self.create_context_menu()
        self.load_stocks()

        # 预设排序设置
        self.sort_column = "change_percent"
        self.sort_reverse = True
        self.refresh_data()  # 首次加载时应用排序


   

    def create_widgets(self):
        columns = ("symbol", "price", "change_percent")
        self.tree = ttk.Treeview(
            self, columns=columns, show="headings",
            selectmode="browse", style="Custom.Treeview"
        )
        
        col_widths = [67, 58, 70]  # 調整後寬度
        headers = ["股票代碼", "新價格", "漲跌幅"]  # 簡化後標題
        
        for col, width, header in zip(columns, col_widths, headers):
            self.tree.heading(col, text=header, 
                            command=lambda c=col: self.treeview_sort_column(c))
            self.tree.column(col, width=width, anchor=tk.CENTER)

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        #ttk.Label(self, text=f"窗格位置：{self.pane_side.upper()}", font=('微軟正黑體', 9)).pack(side=tk.TOP, anchor=tk.W)

    def create_context_menu(self):
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="打開Yahoo Finance頁面", command=self.open_yahoo_finance)
        self.context_menu.add_separator()
       
        move_label = "移動到右窗格" if self.pane_side == "left" else "移動到左窗格"
        self.context_menu.add_command(
            label=move_label,
            command=self.move_to_other_pane
        )
        self.context_menu.add_command(label="複製股票代碼", command=self.copy_symbol)
        self.tree.bind("<Button-3>", self.show_context_menu)

    def move_to_other_pane(self):
        selected = self.tree.selection()
        if not selected:
            return
        
        symbol_item = self.tree.item(selected[0])
        symbol = symbol_item['values'][0]
        
        # 獲取目標窗格
        target_side = "right" if self.pane_side == "left" else "left"
        target_tab = self.main_app.get_current_tab(target_side)
        
        if not target_tab:
            messagebox.showwarning("警告", f"請先在{target_side}窗格建立分頁")
            return
        
        if symbol in target_tab.stocks:
            messagebox.showwarning("警告", f"{symbol} 已存在於目標分頁")
            return
        
        # 執行移動操作
        try:
            # 從當前分頁移除
            self.stocks.remove(symbol)
            self.save_stocks()
            self.refresh_data()
            
            # 添加到目標分頁
            target_tab.stocks.append(symbol)
            target_tab.save_stocks()
            target_tab.refresh_data()
            
            self.main_app.status.config(text=f"已移動 {symbol} 到{target_side}窗格")
        except Exception as e:
            messagebox.showerror("錯誤", f"移動失敗: {str(e)}")

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            try:
                self.context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.context_menu.grab_release()

    def open_yahoo_finance(self):
        selected = self.tree.selection()
        if not selected:
            return
        symbol = self.tree.item(selected[0])['values'][0]
        try:
            webbrowser.open_new_tab(f"https://finance.yahoo.com/chart/{symbol}")
        except Exception as e:
            messagebox.showerror("錯誤", f"無法開啟網頁：{str(e)}")

    def copy_symbol(self):
        selected = self.tree.selection()
        if selected:
            symbol = self.tree.item(selected[0])['values'][0]
            self.clipboard_clear()
            self.clipboard_append(symbol)

    def treeview_sort_column(self, col):
        # 排序时标记缓存失效
        self.cache_valid = False
        if self.sort_column == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = col
            self.sort_reverse = False
            
        for c in self.tree["columns"]:
            heading_text = self.tree.heading(c)["text"]
            if c == col:
                arrow = " ↓" if self.sort_reverse else " ↑"
                self.tree.heading(c, text=heading_text.split(" ↑")[0].split(" ↓")[0] + arrow)
            else:
                self.tree.heading(c, text=heading_text.split(" ↑")[0].split(" ↓")[0])
        
        self.refresh_data()
   

    def load_stocks(self):
        self.stocks = []
        if os.path.exists(self.filename):
            try:
                with open(self.filename, "r") as f:
                    self.stocks = [line.strip() for line in f if line.strip()]
            except Exception as e:
                messagebox.showerror("錯誤", f"讀取失敗：{str(e)}")

    def save_stocks(self):
        try:
            with open(self.filename, "w") as f:
                f.write("\n".join(self.stocks))
        except Exception as e:
            messagebox.showerror("錯誤", f"保存失敗：{str(e)}")

    def delete_stock(self):
        try:
            selected = self.tree.selection()[0]
            symbol = self.tree.item(selected)['values'][0]
            if messagebox.askyesno("確認", f"刪除 {symbol}？"):
                self.stocks.remove(symbol)
                self.save_stocks()
                self.refresh_data()
                return True
        except IndexError:
            messagebox.showwarning("警告", "請選擇股票")
        return False

    def refresh_data(self, force=False):
        """优化后的数据刷新方法"""
        if not force and self.cache_valid:
            return  # 使用缓存数据

        def fetch_data():
            with self.data_lock:
                stocks = self.load_stocks()
                new_items = []
                for symbol in stocks:
                    if symbol not in self.cached_data or force:
                        data = self.get_stock_data(symbol)
                        self.cached_data[symbol] = data
                    else:
                        data = self.cached_data[symbol]
                    # ... 数据处理逻辑 ...
                
                self.after(0, self._update_ui)
                # 在数据获取完成后更新缓存状态
                self.cache_valid = True
                self.cached_items = items  # 存储排序后的数据

        threading.Thread(target=fetch_data, daemon=True).start()

    def _update_ui(self):
        """最小化UI更新操作"""
        self.tree.delete(*self.tree.get_children())
        # 仅插入可见范围数据（分页加载优化）
        for item in self.cached_data.values()[:100]:  # 假设每页显示100条
            self.tree.insert("", "end", values=item)
        # 异步加载剩余数据
        threading.Thread(target=self._load_remaining_data, daemon=True).start()

    def _load_remaining_data(self):
        """后台加载剩余数据"""
        for item in self.cached_data.values()[100:]:
            self.tree.insert("", "end", values=item)

    def refresh_data(self):
        self.tree.delete(*self.tree.get_children())
        items = []
        for symbol in self.stocks:
            data = self.get_stock_data(symbol)
            if data:
                # 只保留需要的數據
                price = f"{data['price']:.1f}" if isinstance(data['price'], float) else 'N/A'
                change_percent = data['change_percent']
                
                item = (
                    data['symbol'],
                    price,
                    change_percent
                )
                items.append((item, data.get('change')))  # 保留change用於排序判斷
        
        # 調整排序邏輯
        if self.sort_column:
            col_index = self.tree["columns"].index(self.sort_column)
            reverse = self.sort_reverse
            if self.sort_column == "change_percent":
                items.sort(key=lambda x: self.parse_percent(x[0][col_index]), reverse=reverse)
           
            elif self.sort_column == "price":
                items.sort(key=lambda x: self.parse_price(x[0][col_index]), reverse=reverse)
            else:
                items.sort(key=lambda x: x[0][col_index], reverse=reverse)

        # 插入簡化後數據
        for item, change in items:
            tags = ()
            try:
                # 解析漲跌幅百分比數值
                percent_str = item[2].replace('%', '')  # 移除百分比符號
                percent = float(percent_str)
                
                # 僅當漲幅超過0.5%時標記為綠色
                if percent > 0.5:
                    tags = ('rise',)
                elif percent < -0.5:  # 新增负值判断
                    tags = ('fall',)
                else:  # 新增中间范围
                    tags = ('neutral',)

            except (ValueError, AttributeError):
                # 異常處理保持原邏輯
               if isinstance(change, float):
                    tags = ('rise',) if change >= 0 else ('fall',)
            
            self.tree.insert("", tk.END, values=item, tags=tags)
        
        # 新增中性样式配置
        self.tree.tag_configure('neutral', foreground='white')
        self.tree.tag_configure('rise', foreground='#33FF77')  #green
        self.tree.tag_configure('fall', foreground='#FF1919')  #red

    def parse_percent(self, value):
        try:
            return float(value.strip('%'))
        except:
            return -float('inf') if self.sort_reverse else float('inf')

    def parse_price(self, value):
        try:
            return float(value)
        except:
            return -float('inf') if self.sort_reverse else float('inf')

    def get_stock_data(self, symbol):
        try:
            ticker = yf.Ticker(symbol)
            data = ticker.info
            market_state = data.get('marketState', 'REGULAR')
            
            price = data.get('regularMarketPrice')
            prev_close = data.get('regularMarketPreviousClose')
            
            change = None
            change_percent = None
            if price and prev_close:
                change = round(price - prev_close, 2)
                change_percent = f"{(change/prev_close)*100:+.2f}%"   ####
            
            
            return {
                "symbol": symbol,
                "price": price,
                "change": change,  # 保留用於顏色標記
                "change_percent": change_percent or 'N/A'
                # 移除其他不需要的欄位
            }
        except Exception as e:
            print(f"獲取數據失敗：{symbol} - {str(e)}")
            return None

class DualPaneStockApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("雙窗看股系統 v2.0")
        self.geometry("428x840")    #####
        self.panes = {"left": {"notebook": None, "tabs": {}}, "right": {"notebook": None, "tabs": {}}}
        self.create_widgets()

        # 配置黑色主题
        self.configure(background='black')
        self._setup_dark_theme()
        self.current_visible_tabs = {}  # 跟踪可见分页

        # 延遲加載配置和自動刷新
        self.after(100, self.initialize_app)

    def initialize_app(self):
        """延遲初始化非必要資源"""
        self.load_config()
        self.after(REFRESH_INTERVAL, self.auto_refresh)

        

    def add_existing_tab(self, side, filename, tab_name):
        """空分頁初始化，延遲加載數據"""
        new_tab = PortfolioTab(self.panes[side]["notebook"], filename, side, self)
        self.panes[side]["notebook"].add(new_tab, text=tab_name)
        self.panes[side]["tabs"][filename] = new_tab
        new_tab.load_stocks(deferred=True)  # 新增延遲加載參數

        


    def _setup_dark_theme(self):
        style = ttk.Style()
        style.theme_use('alt')

        # 全局颜色配置
        bg_color = 'black'
        fg_color = '#ffffff'
        field_bg = '#1f1e1e'
        selected_color = '#1a3d5d'

        style.configure('.', background=bg_color, foreground=fg_color)
        self['bg'] = bg_color

        # 树状视图样式
        style.configure("Custom.Treeview",
                        background=bg_color,
                        fieldbackground=field_bg,
                        foreground=fg_color,
                        rowheight=24,                ####
                        font=('微軟正黑體', 13, 'bold')
                        )
        style.configure("Custom.Treeview.Heading",
                        background=field_bg,
                        foreground=fg_color,
                        font=('微軟正黑體', 10, 'bold')
                        )
        style.map("Custom.Treeview",
                  background=[('selected', selected_color)])

        # 新增滾動條樣式配置
        style.element_create("Vertical.Scrollbar.trough", "from", "clam")
        style.element_create("Vertical.Scrollbar.thumb", "from", "clam")
        style.layout("Vertical.TScrollbar",
            [('Vertical.Scrollbar.trough', {'children':
                [('Vertical.Scrollbar.thumb', {'expand': '1', 'sticky': 'nswe'})]})]
        )
        style.configure("Vertical.TScrollbar",
            troughcolor="#3d3d3d",  # 滑槽顏色
            gripcount=0,
            arrowsize=14,
            darkcolor="grey",    # 滑塊深色部分
            lightcolor="grey",   # 滑塊亮色部分
            background="grey",   # 箭頭顏色
            bordercolor="#2d2d2d",
            arrowcolor="black"
        )

        # 新增分页操作按钮样式
        style.configure('Tab.TButton',
                       background='#2d48e0',  # 蓝色
                       foreground='white',
                       borderwidth=1)
        style.map('Tab.TButton',
                 background=[('active', '#546eff')],
                 foreground=[('active', 'white')])

        # 新增股票操作按钮样式
        style.configure('Stock.TButton',
                       background='#237545',  # 绿色
                       foreground='white',
                       borderwidth=1)
        style.map('Stock.TButton',
                 background=[('active', '#00b34a')],
                 foreground=[('active', 'white')])
        
        # 新增Radiobutton樣式設定
        style.configure('TRadiobutton',
                        background='#2d2d2d',
                        foreground='#ffffff',
                        indicatormargin=4,
                        indicatordiameter=12)
        style.map('TRadiobutton',
                background=[('active', '#3d3d3d')],
                foreground=[('active', '#00ff00')],  # 滑鼠懸停時文字變亮綠色
                indicatorcolor=[
                    ('selected', '#00ff00'),  # 選中時指示器顏色
                    ('!selected', '#555555')  # 未選中時顏色
                ])
        # 更新現有Radiobutton實例
        for rb in self.winfo_children():
            if isinstance(rb, ttk.Radiobutton):
                rb.configure(style='TRadiobutton')

        # 按钮样式
        style.configure('TButton',
                        background=field_bg,
                        foreground=fg_color,
                        borderwidth=1)
        style.map('TButton',
                  background=[('active', selected_color)],
                  foreground=[('active', fg_color)])
        # 标签样式
        style.configure('TLabel',
                        background=bg_color,
                        foreground=fg_color)

        # 分页标签样式
        style.configure('TNotebook.Tab',
                        background=field_bg,
                        foreground=fg_color)
        style.map('TNotebook.Tab',
                  background=[('selected', bg_color)],
                  foreground=[('selected', fg_color)])

        # 状态栏样式
        self.status.configure(style='TLabel')

        # 分隔线样式
        self.option_add('*TCombobox*Listbox.background', field_bg)
        self.option_add('*TCombobox*Listbox.foreground', fg_color)
        self.option_add('*TCombobox*Listbox.selectBackground', selected_color)

    def create_widgets(self):
        main_pane = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True)
 
        # 左窗格
        left_frame = ttk.Frame(main_pane)
        self.create_pane(left_frame, "left")
        main_pane.add(left_frame, weight=1)

        # 右窗格
        right_frame = ttk.Frame(main_pane)
        self.create_pane(right_frame, "right")
        main_pane.add(right_frame, weight=1)

        # 工具列
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X, padx=5, pady=2)
        
        self.side_var = tk.StringVar(value="left")
        ttk.Radiobutton(toolbar, text="左側", variable=self.side_var, value="left").pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(toolbar, text="右側", variable=self.side_var, value="right").pack(side=tk.LEFT, padx=2)
       
        ttk.Button(toolbar, text="＋ 股票", 
                  command=self.add_stock,width=7,
                  style='Stock.TButton').pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="－ 股票", 
                  command=self.delete_stock,width=7,
                  style='Stock.TButton').pack(side=tk.LEFT, padx=2)

        ttk.Button(toolbar, text="＋ 分頁", command=self.add_portfolio,width=7,style='Tab.TButton').pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="－ 分頁", command=self.delete_portfolio,width=7,style='Tab.TButton').pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="↻刷新" , command=self.refresh_all).pack(side=tk.RIGHT, padx=2)

        self.status = ttk.Label(self, text="就緒", anchor=tk.W)
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

        '''style = ttk.Style()
        style.configure("Custom.Treeview", rowheight=25, font=('微軟正黑體', 10))
        style.configure("Custom.Treeview.Heading", font=('微軟正黑體', 10, 'bold'))'''

    def create_pane(self, parent, side):
        pane_frame = ttk.Frame(parent)
        pane_frame.pack(fill=tk.BOTH, expand=True)
        notebook = ttk.Notebook(pane_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        notebook.bind("<<NotebookTabChanged>>", lambda e: self.on_tab_changed(side))
        self.panes[side]["notebook"] = notebook
        self.panes[side]["tabs"] = {}
        ttk.Label(pane_frame, text=f"{side.upper()} 窗格", font=('微軟正黑體', 10, 'bold')).pack(side=tk.TOP, fill=tk.X)

    def add_portfolio(self):
        side = self.side_var.get()
        filename = simpledialog.askstring("新增分頁", f"輸入{side}窗格檔案名稱：")
        if not filename:
            return
        if not filename.endswith(".txt"):
            filename += ".txt"
        
        if filename in self.panes[side]["tabs"]:
            messagebox.showwarning("警告", "分頁已存在")
            return

        new_tab = PortfolioTab(self.panes[side]["notebook"], filename, side, self)
        tab_name = os.path.splitext(filename)[0]
        self.panes[side]["notebook"].add(new_tab, text=tab_name)
        self.panes[side]["tabs"][filename] = new_tab
        self.save_config()

    def delete_portfolio(self):
        side = self.side_var.get()
        current_tab = self.get_current_tab(side)
        if not current_tab:
            return
        
        if messagebox.askyesno("確認", f"刪除分頁 {self.panes[side]['notebook'].tab('current')['text']}？"):
            self.panes[side]["notebook"].forget(current_tab)
            del self.panes[side]["tabs"][current_tab.filename]
            os.remove(current_tab.filename)
            self.save_config()

    def get_current_tab(self, side=None):
        side = side or self.side_var.get()
        notebook = self.panes[side]["notebook"]
        try:
            return notebook.nametowidget(notebook.select())
        except:
            return None

    def on_tab_changed(self, side):
        if current_tab := self.get_current_tab(side):
            current_tab.refresh_data()


    def _update_visible_tabs(self, active_side):
        """资源管理优化"""
        for side in ['left', 'right']:
            if side != active_side:
                tab = self.get_current_tab(side)
                if tab and tab.winfo_ismapped():
                    tab.tree.unmap()  # 隐藏不可见分页组件

    def add_stock(self):
        side = self.side_var.get()
        current_tab = self.get_current_tab(side)
        if not current_tab:
            messagebox.showwarning("警告", "請先選擇分頁")
            return
        
        symbol = simpledialog.askstring("新增股票", "輸入股票代碼：") or ""
        symbol = symbol.upper().strip()
        if not symbol:
            return
        
        if symbol in current_tab.stocks:
            messagebox.showwarning("警告", "股票已存在")
            return
        
        self.status.config(text="驗證中...")
        threading.Thread(target=self.validate_add_stock, args=(side, symbol)).start()

    def validate_add_stock(self, side, symbol):
        try:
            current_tab = self.get_current_tab(side)
            if not current_tab:
                return
            
            data = yf.Ticker(symbol).info
            if not data.get('symbol'):
                raise ValueError("無效代碼")
            
            current_tab.stocks.append(symbol)
            current_tab.save_stocks()
            current_tab.refresh_data()
            self.status.config(text=f"已添加：{symbol}")
        except Exception as e:
            messagebox.showerror("錯誤", f"添加失敗：{str(e)}")
        finally:
            self.status.config(text="就緒")

    def delete_stock(self):
        side = self.side_var.get()
        if current_tab := self.get_current_tab(side):
            if current_tab.delete_stock():
                self.status.config(text="股票已刪除")

    def refresh_all(self):
        for side in ["left", "right"]:
            for tab in self.panes[side]["tabs"].values():
                tab.refresh_data()
        self.status.config(text="全部數據已刷新")

    def auto_refresh(self):
        self.refresh_all()
        self.after(REFRESH_INTERVAL, self.auto_refresh)

    def load_config(self):
        if not os.path.exists(CONFIG_FILE):
            return
        
        try:
            with open(CONFIG_FILE, "r") as f:
                config = json.load(f)
                for side in ["left", "right"]:
                    for filename, tab_name in config.get(side, {}).items():
                        self.add_existing_tab(side, filename, tab_name)
        except Exception as e:
            messagebox.showerror("錯誤", f"配置讀取失敗：{str(e)}")

    def add_existing_tab(self, side, filename, tab_name):
        new_tab = PortfolioTab(self.panes[side]["notebook"], filename, side, self)
        self.panes[side]["notebook"].add(new_tab, text=tab_name)
        self.panes[side]["tabs"][filename] = new_tab

    def save_config(self):
        config = {
            "left": {tab.filename: self.panes["left"]["notebook"].tab(tab, "text") for tab in self.panes["left"]["tabs"].values()},
            "right": {tab.filename: self.panes["right"]["notebook"].tab(tab, "text") for tab in self.panes["right"]["tabs"].values()}
        }
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(config, f, ensure_ascii=False)
        except Exception as e:
            messagebox.showerror("錯誤", f"配置保存失敗：{str(e)}")

if __name__ == "__main__":
    try:
        import yfinance
    except ImportError:
        print("請先安裝套件：pip install yfinance")
        exit()
    
    app = DualPaneStockApp()
    app.mainloop()
