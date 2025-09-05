import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
import cv2
import numpy as np
from PIL import Image, ImageTk
import win32gui
import win32ui
import win32con
import win32api
from ctypes import windll
import os
import json
import csv
from datetime import datetime

class WindowCaptureBot:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("E7 PC FULL AUTO v2.5")
        self.root.geometry("1000x800")  # 增加高度以容納新功能
        
        # 狀態變數
        self.target_window = None
        self.target_hwnd = None
        self.is_running = False
        self.capture_thread = None
        self.match_threshold = 0.8
        
        # ✅ 時間記錄變數
        self.start_time = None
        self.end_time = None
        
        # ✅ 自動次數相關變數
        self.auto_max_count = None  # 最大自動次數（None表示無限）
        self.auto_current_count = 0  # 當前已執行次數
        
        # ✅ 統計變數
        self.stats = {
            'skystones_consumed': 0,      # 天空石消耗
            'covenant_bookmarks': 0,      # 聖約書簽
            'mystic_bookmarks': 0,        # 神秘書簽
            'friendship_bookmarks': 0,    # 友情書簽
            'gold_consumed': 0            # 金幣消耗
        }
        
        # 模板圖片相關
        self.template_images = ["covenant.png", "mystic.png", "friend.png", "text_11.png" , "text_01.png"]
        self.template_vars = []
        self.template_photoimgs = []
        self.loaded_templates = []
        
        # ✅ UI控件列表（用於disable/enable）
        self.ui_controls = []
        
        # 點擊座標設定
        self.click_positions = {
            'scroll_x': 480,  # 640 * 0.75
            'scroll_y': 180,  # 360 / 2
            'left_bottom_x': 107,
            'left_bottom_y': 320,
            'center_confirm_x': 363,
            'center_confirm_y': 256,
            'next_confirm_x': 363,
            'next_confirm_y': 230,
            'cancel_x': 235,
            'cancel_y': 260,
            'scroll_distance': 150
        }
        
        # 設置UI
        self.setup_ui()
        
        # 載入模板圖片
        self.load_template_images()
        
        # 載入設定
        self.load_settings()
        
    def setup_ui(self):
        """設置用戶界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 視窗選擇區域
        window_frame = ttk.LabelFrame(main_frame, text="視窗選擇", padding="5")
        window_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        refresh_btn = ttk.Button(window_frame, text="刷新視窗列表", command=self.refresh_windows)
        refresh_btn.grid(row=0, column=0, padx=(0, 10))
        self.ui_controls.append(refresh_btn)
        
        self.window_var = tk.StringVar()
        self.window_combo = ttk.Combobox(window_frame, textvariable=self.window_var, width=50, state="readonly")
        self.window_combo.grid(row=0, column=1, sticky=(tk.W, tk.E))
        self.window_combo.bind('<<ComboboxSelected>>', self.on_window_selected)
        self.ui_controls.append(self.window_combo)
        
        # 模板圖片多選區域
        template_frame = ttk.LabelFrame(main_frame, text="模板圖片選擇", padding="5")
        template_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 創建模板圖片多選UI
        self.setup_template_selection(template_frame)
        
        # ✅ 統計顯示區域
        stats_frame = ttk.LabelFrame(main_frame, text="統計資訊", padding="5")
        stats_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        self.setup_statistics_display(stats_frame)
        
        # ✅ 自動次數設定區域
        auto_count_frame = ttk.LabelFrame(main_frame, text="自動次數設定", padding="5")
        auto_count_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        self.setup_auto_count_ui(auto_count_frame)
        
        # 匹配閾值設定
        threshold_frame = ttk.LabelFrame(main_frame, text="匹配設定", padding="5")
        threshold_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(threshold_frame, text="匹配閾值:").grid(row=0, column=0, padx=(0, 5))
        self.threshold_var = tk.StringVar(value="0.8")
        threshold_entry = ttk.Entry(threshold_frame, textvariable=self.threshold_var, width=10)
        threshold_entry.grid(row=0, column=1, padx=(0, 20))
        self.ui_controls.append(threshold_entry)
        
        # 控制按鈕
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=5, column=0, columnspan=2, pady=(0, 10))
        
        self.start_button = ttk.Button(control_frame, text="開始自動化", command=self.start_capture)
        self.start_button.grid(row=0, column=0, padx=(0, 10))
        
        self.stop_button = ttk.Button(control_frame, text="停止自動化", command=self.stop_capture, state="disabled")
        self.stop_button.grid(row=0, column=1, padx=(0, 10))
        
        test_btn = ttk.Button(control_frame, text="測試捕獲", command=self.test_capture)
        test_btn.grid(row=0, column=2, padx=(0, 10))
        self.ui_controls.append(test_btn)
        
        save_btn = ttk.Button(control_frame, text="保存設定", command=self.save_settings)
        save_btn.grid(row=0, column=3, padx=(0, 10))
        self.ui_controls.append(save_btn)
        
        # ✅ 重置統計按鈕
        reset_btn = ttk.Button(control_frame, text="重置統計", command=self.reset_statistics)
        reset_btn.grid(row=0, column=4, padx=(0, 10))
        self.ui_controls.append(reset_btn)
        
        # 狀態顯示
        status_frame = ttk.LabelFrame(main_frame, text="狀態資訊", padding="5")
        status_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        self.status_text = tk.Text(status_frame, height=15, width=90)
        scrollbar = ttk.Scrollbar(status_frame, orient="vertical", command=self.status_text.yview)
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        # 不要將 status_text 加入 ui_controls
        
        # 配置網格權重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(6, weight=1)
        window_frame.columnconfigure(1, weight=1)
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(0, weight=1)
    
    def setup_template_selection(self, parent_frame):
        """設置模板圖片選擇UI - 排除文字模板"""
        # ✅ 只為前3個模板(covenant, mystic, friend)創建UI
        display_templates = self.template_images[:3]  # 只顯示前3個
        
        for i, img_name in enumerate(display_templates):
            # 創建每個圖片的框架
            img_frame = ttk.Frame(parent_frame)
            img_frame.grid(row=i//3, column=i%3, padx=10, pady=5, sticky=(tk.W, tk.E))
            
            # 載入並顯示縮圖
            img_path = os.path.join("image", img_name)
            if os.path.exists(img_path):
                try:
                    pil_img = Image.open(img_path)
                    pil_img = pil_img.resize((80, 80), Image.Resampling.LANCZOS)
                    photo_img = ImageTk.PhotoImage(pil_img)
                    self.template_photoimgs.append(photo_img)
                    
                    # 圖片標籤
                    img_label = ttk.Label(img_frame, image=photo_img)
                    img_label.grid(row=0, column=0, padx=(0, 5))
                    
                except Exception as e:
                    self.log_message(f"載入圖片 {img_name} 失敗: {e}")
                    placeholder = ttk.Label(img_frame, text="圖片\n載入\n失敗", width=10)
                    placeholder.grid(row=0, column=0, padx=(0, 5))
                    self.template_photoimgs.append(None)
            else:
                placeholder = ttk.Label(img_frame, text="圖片\n未找到", width=10)
                placeholder.grid(row=0, column=0, padx=(0, 5))
                self.template_photoimgs.append(None)
            
            # 複選框 - friend.png 默認不勾選
            default_value = True if img_name != "friend.png" else False
            var = tk.BooleanVar(value=default_value)
            checkbox = ttk.Checkbutton(img_frame, text=img_name.replace('.png', ''), variable=var)
            checkbox.grid(row=1, column=0)
            
            self.template_vars.append(var)
            self.ui_controls.append(checkbox)

    def setup_statistics_display(self, parent_frame):
        """✅ 設置統計顯示UI"""
        # 第一行：天空石和金幣
        row1_frame = ttk.Frame(parent_frame)
        row1_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 5))
        
        ttk.Label(row1_frame, text="天空石消耗:", font=("Arial", 10)).grid(row=0, column=0, padx=(0, 5))
        self.skystone_label = ttk.Label(row1_frame, text="0", font=("Arial", 10, "bold"), foreground="purple")
        self.skystone_label.grid(row=0, column=1, padx=(0, 20))
        
        ttk.Label(row1_frame, text="金幣消耗:", font=("Arial", 10)).grid(row=0, column=2, padx=(0, 5))
        self.gold_label = ttk.Label(row1_frame, text="0", font=("Arial", 10, "bold"), foreground="orange")
        self.gold_label.grid(row=0, column=3, padx=(0, 20))
        
        # 第二行：三種書簽
        row2_frame = ttk.Frame(parent_frame)
        row2_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
        
        ttk.Label(row2_frame, text="聖約書簽:", font=("Arial", 10)).grid(row=0, column=0, padx=(0, 5))
        self.covenant_label = ttk.Label(row2_frame, text="0", font=("Arial", 10, "bold"), foreground="blue")
        self.covenant_label.grid(row=0, column=1, padx=(0, 15))
        
        ttk.Label(row2_frame, text="神秘書簽:", font=("Arial", 10)).grid(row=0, column=2, padx=(0, 5))
        self.mystic_label = ttk.Label(row2_frame, text="0", font=("Arial", 10, "bold"), foreground="red")
        self.mystic_label.grid(row=0, column=3, padx=(0, 15))
        
        ttk.Label(row2_frame, text="友情書簽:", font=("Arial", 10)).grid(row=0, column=4, padx=(0, 5))
        self.friendship_label = ttk.Label(row2_frame, text="0", font=("Arial", 10, "bold"), foreground="green")
        self.friendship_label.grid(row=0, column=5)
    
    def setup_auto_count_ui(self, parent_frame):
        """✅ 設置自動次數UI"""
        # 自動次數設定
        ttk.Label(parent_frame, text="自動次數 (空白=無限):", font=("Arial", 10)).grid(row=0, column=0, padx=(0, 5))
        
        self.auto_count_var = tk.StringVar()
        auto_count_entry = ttk.Entry(parent_frame, textvariable=self.auto_count_var, width=10)
        auto_count_entry.grid(row=0, column=1, padx=(0, 20))
        self.ui_controls.append(auto_count_entry)
        
        # 已刷新次數顯示
        ttk.Label(parent_frame, text="已刷新次數:", font=("Arial", 10)).grid(row=0, column=2, padx=(20, 5))
        self.current_count_label = ttk.Label(parent_frame, text="0", font=("Arial", 10, "bold"), foreground="darkgreen")
        self.current_count_label.grid(row=0, column=3, padx=(0, 10))
        
        # 進度顯示
        self.progress_label = ttk.Label(parent_frame, text="", font=("Arial", 9), foreground="gray")
        self.progress_label.grid(row=0, column=4, padx=(10, 0))
    
    def update_auto_count_display(self):
        """✅ 更新自動次數顯示"""
        self.current_count_label.config(text=f"{self.auto_current_count}")
        
        # 更新進度顯示
        if self.auto_max_count is not None:
            progress_text = f"({self.auto_current_count}/{self.auto_max_count})"
            self.progress_label.config(text=progress_text)
        else:
            self.progress_label.config(text="(無限)")
    
    def update_statistics_display(self):
        """✅ 更新統計顯示"""
        self.skystone_label.config(text=f"{self.stats['skystones_consumed']:,}")
        self.gold_label.config(text=f"{self.stats['gold_consumed']:,}")
        self.covenant_label.config(text=f"{self.stats['covenant_bookmarks']:,}")
        self.mystic_label.config(text=f"{self.stats['mystic_bookmarks']:,}")
        self.friendship_label.config(text=f"{self.stats['friendship_bookmarks']:,}")
    
    def reset_statistics(self):
        """✅ 重置統計數據"""
        self.stats = {
            'skystones_consumed': 0,
            'covenant_bookmarks': 0,
            'mystic_bookmarks': 0,
            'friendship_bookmarks': 0,
            'gold_consumed': 0
        }
        # ✅ 同時重置自動次數
        self.auto_current_count = 0
        self.update_statistics_display()
        self.update_auto_count_display()
        self.log_message("統計數據和計數已重置", color="gray")
    
    def toggle_ui_controls(self, enabled):
        """✅ 啟用/禁用UI控件 - 修正版"""
        state = "normal" if enabled else "disabled"
        for control in self.ui_controls:
            try:
                control.config(state=state)
            except:
                pass
    
    def write_summary_to_csv(self, duration_seconds):
        """✅ 將自動化總結寫入CSV文件"""
        try:
            filename = "automation_summary.csv"
            file_exists = os.path.isfile(filename)
            
            # CSV欄位定義
            fieldnames = [
                '開始時間', '結束時間', '使用時間(HH:MM:SS)',
                '刷新次數', '天空石消耗', '聖約書籤獲得', '神秘書籤獲得',
                '友情書籤獲得', '金幣消耗'
            ]
            
            with open(filename, mode='a', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                
                # 如果文件不存在，寫入標題行
                if not file_exists:
                    writer.writeheader()
                
                # 格式化時間
                start_str = datetime.fromtimestamp(self.start_time).strftime('%Y-%m-%d %H:%M:%S')
                end_str = datetime.fromtimestamp(self.end_time).strftime('%Y-%m-%d %H:%M:%S')
                
                # 格式化持續時間
                hours = int(duration_seconds // 3600)
                minutes = int((duration_seconds % 3600) // 60)
                seconds = int(duration_seconds % 60)
                duration_formatted = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                
                # 寫入數據行
                row = {
                    '開始時間': start_str,
                    '結束時間': end_str,
                    '使用時間(HH:MM:SS)': duration_formatted,
                    '刷新次數': self.auto_current_count,
                    '天空石消耗': self.stats['skystones_consumed'],
                    '聖約書籤獲得': self.stats['covenant_bookmarks'],
                    '神秘書籤獲得': self.stats['mystic_bookmarks'],
                    '友情書籤獲得': self.stats['friendship_bookmarks'],
                    '金幣消耗': self.stats['gold_consumed']
                }
                
                writer.writerow(row)
            
            self.log_message(f"📊 自動化總結已匯出至 {filename}", color="green")
            
        except Exception as e:
            self.log_message(f"匯出CSV時發生錯誤: {e}", color="red")

    # 以下是其他原有方法，保持不變...
    def load_template_images(self):
        """載入模板圖片到記憶體 - 包含文字模板"""
        self.loaded_templates = []
        for img_name in self.template_images:
            img_path = os.path.join("image", img_name)
            if os.path.exists(img_path):
                try:
                    # ✅ 直接載入為灰階圖片
                    template = cv2.imread(img_path, cv2.IMREAD_GRAYSCALE)
                    if template is not None:
                        self.loaded_templates.append(template)
                        if img_name.startswith("text"):
                            self.log_message(f"載入文字模板: {img_name}")
                        else:
                            self.log_message(f"載入模板圖片: {img_name}")
                    else:
                        self.loaded_templates.append(None)
                        self.log_message(f"載入模板圖片失敗: {img_name}")
                except Exception as e:
                    self.loaded_templates.append(None)
                    self.log_message(f"載入模板圖片錯誤 {img_name}: {e}")
            else:
                self.loaded_templates.append(None)
                self.log_message(f"模板圖片不存在: {img_path}")

    def log_message(self, message, color="black"):
        """✅ 記錄訊息到狀態文本框（支援顏色）"""
        timestamp = time.strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        # 插入文字
        self.status_text.insert(tk.END, log_entry)
        
        # 如果有顏色設定，應用顏色
        if color != "black":
            # 獲取剛插入文字的起始和結束位置
            start_line = int(self.status_text.index(tk.END).split('.')[0]) - 2
            start_pos = f"{start_line}.0"
            end_pos = f"{start_line}.end"
            
            # 創建標籤並應用顏色
            tag_name = f"color_{color}_{time.time()}"
            self.status_text.tag_add(tag_name, start_pos, end_pos)
            self.status_text.tag_config(tag_name, foreground=color)
        
        self.status_text.see(tk.END)
        self.root.update_idletasks()

    def refresh_windows(self):
        """刷新可用視窗列表"""
        windows = []
        
        def enum_windows_proc(hwnd, lParam):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                window_text = win32gui.GetWindowText(hwnd)
                if len(window_text) > 3:  # 過濾掉太短的標題
                    windows.append((hwnd, window_text))
            return True
        
        win32gui.EnumWindows(enum_windows_proc, 0)
        
        # 更新下拉選單
        window_titles = [f"{hwnd}: {title}" for hwnd, title in windows]
        self.window_combo['values'] = window_titles
        
        self.log_message(f"找到 {len(windows)} 個視窗")
    
    def on_window_selected(self, event):
        """當選擇視窗時的回調函數"""
        selected = self.window_var.get()
        if selected:
            hwnd = int(selected.split(':')[0])
            self.target_hwnd = hwnd
            self.target_window = win32gui.GetWindowText(hwnd)
            self.log_message(f"選擇目標視窗: {self.target_window}")
    
    def capture_window(self, hwnd):
        """捕獲指定視窗的畫面"""
        try:
            # 獲取視窗大小和位置
            left, top, right, bottom = win32gui.GetClientRect(hwnd)
            width = right - left
            height = bottom - top
            
            # 創建設備上下文
            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()
            
            # 創建位圖對象
            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
            saveDC.SelectObject(saveBitMap)
            
            # 拷貝屏幕內容
            result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 3)
            
            # 轉換為numpy陣列
            bmpinfo = saveBitMap.GetInfo()
            bmpstr = saveBitMap.GetBitmapBits(True)
            
            img = np.frombuffer(bmpstr, dtype='uint8')
            img.shape = (height, width, 4)
            gimg = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            
            # 清理資源
            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(hwnd, hwndDC)
            if result == 1:
                return gimg
            else:
                return None
        except Exception as e:
            self.log_message(f"捕獲視窗時發生錯誤: {e}")
            return None

    def find_template_in_image(self, image, template):
        """在圖片中尋找模板"""
        try:
            if template is None:
                return None, 0
            
            # 使用模板匹配
            result = cv2.matchTemplate(image, template, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            
            if max_val >= self.match_threshold:
                return max_loc, max_val
            else:
                return None, max_val
                
        except Exception as e:
            self.log_message(f"模板匹配時發生錯誤: {e}")
            return None, 0
    
    def click_at_position(self, hwnd, x, y, click_time = 3):
        """使用 PostMessage 連續點擊三次，每次間隔0.1秒"""
        try:
            lParam_start = win32api.MAKELONG(x, y)
            
            # 連續點擊三次
            for i in range(click_time):
                # 移動滑鼠到位置
                win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lParam_start)
                time.sleep(0.05)
                
                # 按下左鍵
                win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam_start)
                time.sleep(0.1)  # 長按效果
                
                # 釋放左鍵
                win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, lParam_start)
                
                # 如果不是最後一次點擊，等待間隔時間
                if i < click_time-1:  # 前兩次點擊後等待
                    time.sleep(0.1)
            
            return True
            
        except Exception as e:
            self.log_message(f"點擊錯誤: {e}")
            return False

    def simulate_vertical_scroll(self, hwnd, start_x, start_y, distance=100):
        """更真實的長按拖動模擬"""
        try:
            # 1. 發送鼠標移動到起始位置
            lParam_start = win32api.MAKELONG(start_x, start_y)
            win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lParam_start)
            time.sleep(0.05)
            
            # 2. 按下左鍵
            win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam_start)
            time.sleep(0.1)  # 長按效果
            
            # 3. 平滑拖動（每20px一步）
            end_y = start_y - distance
            current_y = start_y
            step_size = 20
            
            while current_y > end_y:
                current_y -= step_size
                if current_y < end_y:
                    current_y = end_y
                
                lParam_move = win32api.MAKELONG(start_x, current_y)
                win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, win32con.MK_LBUTTON, lParam_move)
                time.sleep(0.03)  # 控制拖動速度
            
            # 4. 確認到達最終位置
            lParam_end = win32api.MAKELONG(start_x, end_y)
            win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, win32con.MK_LBUTTON, lParam_end)
            time.sleep(0.1)
            
            # 5. 釋放左鍵
            win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, lParam_end)
            win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, 0, lParam_end)
            return True
            
        except Exception as e:
            # 安全釋放左鍵
            try:
                lParam_end = win32api.MAKELONG(start_x, start_y - distance)
                win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, lParam_end)
            except:
                pass
            self.log_message(f"拖動滾動錯誤: {e}")
            return False

    def resize_target_window(self, hwnd):
        """將目標視窗調整為640x360"""
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.3)  # 等待還原完成
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            win32gui.MoveWindow(hwnd, left, top, 640, 360, True)
            self.log_message("目標視窗已調整為640x360")
            return True
        except Exception as e:
            self.log_message(f"調整視窗大小時發生錯誤: {e}")
            return False
    
    def capture_loop(self):
        """主要的自動化循環 - 加入狀態判定"""
        # 調整視窗大小
        if not self.resize_target_window(self.target_hwnd):
            self.log_message("無法調整視窗大小，停止自動化")
            self.stop_capture()
            return
        
        self.log_message("自動化進行時滑鼠不要進入目標視窗，會影響點擊準確度!!!", color="red")
        time.sleep(1)
        
        while self.is_running:
            try:
                # 獲取選中的模板
                selected_templates = []
                selected_names = []
                for i, (var, template) in enumerate(zip(self.template_vars, self.loaded_templates)):
                    if var.get() and template is not None:
                        selected_templates.append(template)
                        selected_names.append(self.template_images[i])
                
                if not selected_templates:
                    self.log_message("沒有選中的模板圖片")
                    time.sleep(1)
                    continue
                
                # 第一次捕獲和檢測
                captured_image = self.capture_window(self.target_hwnd)
                if captured_image is None:
                    self.log_message("無法捕獲視窗畫面")
                    time.sleep(1)
                    continue
                    
                if not self.is_running:
                    break
                    
                # 檢查是否找到匹配的模板
                for template, name in zip(selected_templates, selected_names):
                    match_loc, confidence = self.find_template_in_image(captured_image, template)
                    if match_loc is not None:
                        match_x, match_y = match_loc
                        self.log_message(f"找到匹配圖片 {name} 於位置 ({match_x}, {match_y}), 信心度: {confidence:.3f}")
                        
                        # ✅ 檢查點擊狀態
                        is_clickable, status = self.check_clickable_status(captured_image, match_x, match_y)
                        
                        if is_clickable:
                            # 可以點擊 - 執行原本邏輯
                            self.log_message(f"狀態檢查: {status} - 可以點擊", color="green")
                            
                            # 點擊匹配位置
                            click_x = match_x + 280
                            click_y = match_y + 25
                            time.sleep(1)
                            self.click_at_position(self.target_hwnd, click_x, click_y)
                            time.sleep(1)
                            
                            if not self.is_running:
                                break
                                
                            # 點擊畫面中央確認
                            self.click_at_position(self.target_hwnd, self.click_positions['center_confirm_x'], 
                                                self.click_positions['center_confirm_y'])
                                                        
                            # 根據圖片類型添加統計和顏色日誌
                            if name == "covenant.png":
                                self.stats['covenant_bookmarks'] += 5
                                self.stats['gold_consumed'] += 184000
                                self.log_message("找到聖約書簽！", color="blue")
                            elif name == "mystic.png":
                                self.stats['mystic_bookmarks'] += 50
                                self.stats['gold_consumed'] += 280000
                                self.log_message("找到神秘書簽！", color="red")
                            elif name == "friend.png":
                                self.stats['friendship_bookmarks'] += 5
                                self.stats['gold_consumed'] += 18000
                                self.log_message("找到友情書簽！", color="green")
                            
                            # 更新統計顯示
                            self.update_statistics_display()
                            time.sleep(1)
                        else:
                            # 不能點擊 - 記錄已購買
                            self.log_message(f"狀態檢查: {status} - 已購買，跳過", color="orange")
                
                if not self.is_running:
                    break
                
                # 沒有找到可點擊的目標，執行滑動
                self.simulate_vertical_scroll(
                    self.target_hwnd, 
                    self.click_positions['scroll_x'], 
                    self.click_positions['scroll_y'],
                    self.click_positions['scroll_distance']
                )
                
                time.sleep(2)
                
                # 滑動後再次檢測 (重複相同邏輯)
                captured_image2 = self.capture_window(self.target_hwnd)
                if captured_image2 is None:
                    continue
                    
                if not self.is_running:
                    break
                    
                for template, name in zip(selected_templates, selected_names):
                    match_loc, confidence = self.find_template_in_image(captured_image2, template)
                    if match_loc is not None:
                        match_x, match_y = match_loc
                        self.log_message(f"滑動後找到匹配圖片 {name} 於位置 ({match_x}, {match_y})")
                        
                        # ✅ 檢查點擊狀態
                        is_clickable, status = self.check_clickable_status(captured_image2, match_x, match_y)
                        
                        if is_clickable:
                            # 執行相同的點擊邏輯...
                            self.log_message(f"滑動後狀態檢查: {status} - 可以點擊", color="green")
                            # 點擊匹配位置
                            click_x = match_x + 280
                            click_y = match_y + 25
                            time.sleep(1)
                            self.click_at_position(self.target_hwnd, click_x, click_y)
                            time.sleep(1)
                            
                            if not self.is_running:
                                break
                                
                            # 點擊畫面中央確認
                            self.click_at_position(self.target_hwnd, self.click_positions['center_confirm_x'], 
                                                self.click_positions['center_confirm_y'])
                                                
                            # 根據圖片類型添加統計和顏色日誌
                            if name == "covenant.png":
                                self.stats['covenant_bookmarks'] += 5
                                self.stats['gold_consumed'] += 184000
                                self.log_message("找到聖約書簽！", color="blue")
                            elif name == "mystic.png":
                                self.stats['mystic_bookmarks'] += 50
                                self.stats['gold_consumed'] += 280000
                                self.log_message("找到神秘書簽！", color="red")
                            elif name == "friend.png":
                                self.stats['friendship_bookmarks'] += 5
                                self.stats['gold_consumed'] += 18000
                                self.log_message("找到友情書簽！", color="green")
                            
                            # 更新統計顯示
                            self.update_statistics_display()
                            time.sleep(1)
                            break
                        else:
                            self.log_message(f"滑動後狀態檢查: {status} - 已購買，跳過", color="orange")
                
                self.auto_current_count += 1
                
                # 檢查次數限制
                if self.auto_max_count is not None and self.auto_current_count >= self.auto_max_count:
                    self.log_message(f"✅ 已達到設定的最大次數 {self.auto_max_count}，自動停止", color="green")
                    self.stop_capture()
                    break
                
                # 執行底部點擊流程
                self.click_at_position(self.target_hwnd, self.click_positions['left_bottom_x'], 
                                        self.click_positions['left_bottom_y'])
                time.sleep(1)
                
                if not self.is_running:
                    break
                    
                self.click_at_position(self.target_hwnd, self.click_positions['next_confirm_x'], 
                                        self.click_positions['next_confirm_y'])
                
                # 增加天空石消耗
                if self.is_running:
                    self.stats['skystones_consumed'] += 3
                    self.update_statistics_display()
                    self.update_auto_count_display()
                    
                self.click_at_position(self.target_hwnd, self.click_positions['cancel_x'], 
                        self.click_positions['cancel_y'], 2)
                time.sleep(2)
                    
            except Exception as e:
                self.log_message(f"自動化循環中發生錯誤: {e}")
                time.sleep(1)

    
    def start_capture(self):
        """✅ 開始自動化（記錄開始時間）"""
        if not self.target_hwnd:
            messagebox.showerror("錯誤", "請先選擇目標視窗")
            return
        
        # 檢查是否有選中的模板
        has_selected = any(var.get() for var in self.template_vars)
        if not has_selected:
            messagebox.showerror("錯誤", "請至少選擇一個模板圖片")
            return
        
        try:
            self.match_threshold = float(self.threshold_var.get())
        except ValueError:
            messagebox.showerror("錯誤", "請輸入有效的匹配閾值")
            return
        
        # ✅ 設定自動次數限制
        try:
            auto_count_text = self.auto_count_var.get().strip()
            if auto_count_text == "" or auto_count_text == "0":
                self.auto_max_count = None  # 無限
                self.log_message("設定為無限次數模式")
            else:
                self.auto_max_count = int(auto_count_text)
                self.log_message(f"設定最大自動次數: {self.auto_max_count}")
        except ValueError:
            messagebox.showerror("錯誤", "請輸入有效的自動次數（數字或留空）")
            return
        
        # ✅ 記錄開始時間和重置計數器
        self.start_time = time.time()
        self.auto_current_count = 0
        self.update_auto_count_display()
        
        self.is_running = True
        
        # ✅ 禁用UI控件
        self.toggle_ui_controls(False)
        
        self.capture_thread = threading.Thread(target=self.capture_loop, daemon=True)
        self.capture_thread.start()
        
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        
        if self.auto_max_count is not None:
            self.log_message(f"🚀 開始自動化流程（限制 {self.auto_max_count} 次）...")
        else:
            self.log_message("🚀 開始自動化流程（無限次數）...")
    
    def stop_capture(self):
        """✅ 停止自動化（記錄結束時間並匯出CSV）"""
        self.is_running = False
        
        # ✅ 記錄結束時間
        self.end_time = time.time()
        
        # ✅ 重新啟用UI控件
        self.toggle_ui_controls(True)
        
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        
        # ✅ 計算運行時間並匯出CSV
        if self.start_time is not None:
            duration_seconds = self.end_time - self.start_time
            self.write_summary_to_csv(duration_seconds)
            
            # 格式化顯示運行時間
            hours = int(duration_seconds // 3600)
            minutes = int((duration_seconds % 3600) // 60)
            seconds = int(duration_seconds % 60)
            time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
            
            if self.auto_max_count is not None:
                self.log_message(f"⏹️ 停止自動化流程（已執行 {self.auto_current_count}/{self.auto_max_count} 次，運行時間: {time_str}）")
            else:
                self.log_message(f"⏹️ 停止自動化流程（已執行 {self.auto_current_count} 次，運行時間: {time_str}）")
        else:
            self.log_message(f"⏹️ 停止自動化流程（已執行 {self.auto_current_count} 次）")
    
    def check_clickable_status(self, image, match_x, match_y):
        """使用預載入模板檢查點擊狀態"""
        try:
            check_x = match_x + 255
            check_y = match_y + 22
            
            # 確保區域在圖片範圍內
            h, w = image.shape[:2]
            if check_x + 25 > w or check_y + 13 > h or check_x < 0 or check_y < 0:
                return False, "區域超出範圍"
            
            # 提取ROI區域
            roi_gray = image[check_y:check_y+20, check_x:check_x+30]
            #roi_gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
            
            # ✅ 使用預載入的文字模板
            try:
                text11_idx = self.template_images.index("text_11.png")
                text01_idx = self.template_images.index("text_01.png")
                
                template_11 = self.loaded_templates[text11_idx]
                template_01 = self.loaded_templates[text01_idx]
            except (ValueError, IndexError):
                self.log_message("找不到文字模板索引")
                return False, "模板索引錯誤"
            
            if template_11 is not None and template_01 is not None:
                # 模板匹配
                result_11 = cv2.matchTemplate(roi_gray, template_11, cv2.TM_CCOEFF_NORMED)
                result_01 = cv2.matchTemplate(roi_gray, template_01, cv2.TM_CCOEFF_NORMED)
                
                _, max_val_11, _, _ = cv2.minMaxLoc(result_11)
                _, max_val_01, _, _ = cv2.minMaxLoc(result_01)
                
                self.log_message(f"文字模板匹配: 1/1={max_val_11:.3f}, 0/1={max_val_01:.3f}")
                
                threshold = 0.7  # 匹配閾值
                
                if max_val_11 > threshold and max_val_11 > max_val_01:
                    return True, "1/1"
                elif max_val_01 > threshold:
                    return False, "0/1"
                else:
                    self.log_message(f"文字模板匹配度過低: 1/1={max_val_11:.3f}, 0/1={max_val_01:.3f}")
                    return False, "匹配度不足"
            else:
                missing = []
                if template_11 is None:
                    missing.append("text_11.png")
                if template_01 is None:
                    missing.append("text_01.png")
                self.log_message(f"文字模板未載入: {', '.join(missing)}")
                return False, f"模板未載入: {', '.join(missing)}"
                
        except Exception as e:
            self.log_message(f"文字模板匹配錯誤: {e}")
            return False, "檢查錯誤"

    def draw_debug_rectangle(self, image, match_x, match_y, is_clickable):
        """在圖片上畫出判定區域的紅框 (用於debug)"""
        try:
            check_x = match_x + 255
            check_y = match_y + 22
            
            # 記錄座標信息
            self.log_message(f"準備畫框: 原始位置({match_x}, {match_y}) -> 檢查位置({check_x}, {check_y})")
            
            # 確保座標在圖片範圍內
            h, w = image.shape[:2]
            if check_x < 0 or check_y < 0 or check_x + 20 > w or check_y + 10 > h:
                self.log_message(f"⚠️ 畫框位置超出圖片範圍! 位置:({check_x},{check_y}) 圖片:{w}x{h}")
                # 即使超出範圍，也畫一個小框標示
                check_x = max(0, min(check_x, w-21))
                check_y = max(0, min(check_y, h-11))
            
            # 設置顏色和線寬
            if is_clickable:
                color = (0, 255, 0)  # 綠色 BGR
                status_text = "OK"
            else:
                color = (0, 0, 255)  # 紅色 BGR  
                status_text = "BOUGHT"
            
            # 畫粗一點的矩形框 (線寬3)
            cv2.rectangle(image, (check_x, check_y), (check_x + 25, check_y + 13), color, 3)
            
            # 在框下方加文字標示
            text_y = check_y + 25 if check_y + 25 < h else check_y - 5
            cv2.putText(image, status_text, (check_x, text_y), 
                    cv2.FONT_HERSHEY_SIMPLEX, 0.4, color, 2)
            
            # 同時在模板匹配位置畫一個藍色框
            cv2.rectangle(image, (match_x, match_y), (match_x + 50, match_y + 30), (255, 0, 0), 2)
            cv2.putText(image, "MATCH", (match_x, match_y - 5), 
                    cv2.FONT_HERSHEY_SIMPLEX, 0.4, (255, 0, 0), 2)
            
            self.log_message(f"✅ 成功畫框: 檢查區域({check_x},{check_y}) 顏色:{color} 狀態:{status_text}")
            return image
            
        except Exception as e:
            self.log_message(f"畫框錯誤: {e}")
            return image

    def test_capture(self):
        """測試捕獲功能 - 加入debug紅框"""
        if not self.target_hwnd:
            messagebox.showerror("錯誤", "請先選擇目標視窗")
            return
        
        self.resize_target_window(self.target_hwnd)
        time.sleep(1)
        
        captured_image = self.capture_window(self.target_hwnd)
        if captured_image is not None:
            debug_image = captured_image.copy()
            
            # ✅ 只測試前3個模板（排除文字模板）
            selected_count = 0
            for i in range(3):  # 只處理前3個模板
                var = self.template_vars[i]
                template = self.loaded_templates[i]
                
                if var.get() and template is not None:
                    selected_count += 1
                    match_loc, confidence = self.find_template_in_image(captured_image, template)
                    name = self.template_images[i]
                    
                    if match_loc is not None:
                        match_x, match_y = match_loc
                        self.log_message(f"測試匹配 {name}: 找到於({match_x}, {match_y}), 信心度: {confidence:.3f}")
                        
                        # ✅ 檢查狀態並畫debug框
                        is_clickable, status = self.check_clickable_status(captured_image, match_x, match_y)
                        debug_image = self.draw_debug_rectangle(debug_image, match_x, match_y, is_clickable)
                        
                        self.log_message(f"狀態檢查: {status} - {'可點擊' if is_clickable else '已購買'}")
                    else:
                        self.log_message(f"測試匹配 {name}: 未找到, 最高信心度: {confidence:.3f}")
            
            # 保存原圖和debug圖
            cv2.imwrite("test_capture.png", captured_image)
            cv2.imwrite("test_capture_debug.png", debug_image)
            
            self.log_message("測試捕獲成功，圖片保存為 test_capture.png")
            self.log_message("Debug圖片保存為 test_capture_debug.png (含紅框標記)")
            
            if selected_count == 0:
                self.log_message("未選擇任何模板進行測試")
        else:
            self.log_message("測試捕獲失敗")

        
    def save_settings(self):
        """✅ 保存設定（包含統計數據和自動次數設定）"""
        settings = {
            "window": self.window_var.get(),
            "threshold": self.threshold_var.get(),
            "auto_count": self.auto_count_var.get(),
            "template_selections": [var.get() for var in self.template_vars],
            "statistics": self.stats
        }
        
        try:
            with open("settings.json", "w", encoding="utf-8") as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            self.log_message("設定和統計已保存")
        except Exception as e:
            self.log_message(f"保存設定時發生錯誤: {e}")
    
    def load_settings(self):
        """✅ 載入設定（包含統計數據和自動次數設定）"""
        try:
            if os.path.exists("settings.json"):
                with open("settings.json", "r", encoding="utf-8") as f:
                    settings = json.load(f)
                
                self.threshold_var.set(settings.get("threshold", "0.8"))
                self.auto_count_var.set(settings.get("auto_count", ""))
                
                # 載入模板選擇狀態（friend.png 默認不勾選）
                template_selections = settings.get("template_selections", [True, True, False])
                for i, selection in enumerate(template_selections):
                    if i < len(self.template_vars):
                        self.template_vars[i].set(selection)
                
                # ✅ 載入統計數據
                saved_stats = settings.get("statistics", {})
                for key in self.stats:
                    if key in saved_stats:
                        self.stats[key] = saved_stats[key]
                
                self.update_statistics_display()
                self.update_auto_count_display()
                self.log_message("設定和統計已載入")
        except Exception as e:
            self.log_message(f"載入設定時發生錯誤: {e}")
    
    def run(self):
        """運行應用程序"""
        self.refresh_windows()
        self.log_message("🚀E7 PC FULL AUTO v2.5 已啟動", color="green")
        self.log_message("開始前先確保 Windows顯示設定->縮放與配置->比例 為100%", color="red")
        self.root.mainloop()

if __name__ == "__main__":
    try:
        # 檢查 image 資料夾是否存在
        if not os.path.exists("image"):
            os.makedirs("image")
            print("已創建 image 資料夾，請將 covenant.png, mystic.png, friend.png 放入此資料夾")
        
        app = WindowCaptureBot()
        app.run()
    except Exception as e:
        print(f"程序運行時發生錯誤: {e}")
        input("按 Enter 鍵退出...")
