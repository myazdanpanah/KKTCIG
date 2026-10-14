# service_monitor.py
# مانیتور سرویس تولید فایل‌های نهایی – با قابلیت توقف سرویس و غیرفعال کردن دکمه اجرا هنگام اجرای سرویس

import os
import sys
import json
import subprocess
import threading
import time
import requests
import customtkinter as ctk
from PIL import Image, ImageDraw
import pystray
from pystray import MenuItem as item
from plyer import notification
import psutil  # برای مدیریت فرآیندها

# ==================== تنظیم فونت ====================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

FONT_FAMILY = "B Nazanin"
FONT_SIZE_NORMAL = 14
FONT_SIZE_BOLD = 16
FONT_SIZE_TITLE = 20

# ==================== خواندن تنظیمات ====================
CONFIG_FILE = "monitor_config.json"

def load_config():
    if not os.path.exists(CONFIG_FILE):
        default_config = {
            "client_app_path": "",
            "api_base_url": "http://127.0.0.1:5000",
            "service_process_name": "final_file_service.exe",  # نام فرآیند سرویس (بدون مسیر)
            "refresh_interval": 5,
            "enable_notifications": True
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, indent=2, ensure_ascii=False)
        return default_config
    with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

CONFIG = load_config()
API_BASE_URL = CONFIG.get("api_base_url", "http://127.0.0.1:5000")
REFRESH_INTERVAL = CONFIG.get("refresh_interval", 5)
CLIENT_APP_PATH = CONFIG.get("client_app_path", "")
ENABLE_NOTIFICATIONS = CONFIG.get("enable_notifications", True)
SERVICE_PROCESS_NAME = CONFIG.get("service_process_name", "final_file_service.exe")

class ServiceMonitorApp:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("مانیتور سرویس تولید فایل نهایی")
        self.root.geometry("600x560")
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)
        
        self.service_running = False
        self.last_status = None
        self.icon = None
        self.tray_thread = None
        self.service_pid = None  # برای ذخیره PID سرویس
        
        # ویجت‌ها
        self.status_label = ctk.CTkLabel(
            self.root, text="وضعیت: در حال بررسی...",
            font=ctk.CTkFont(family=FONT_FAMILY, size=FONT_SIZE_BOLD, weight="bold")
        )
        self.status_label.pack(pady=10)
        
        self.files_label = ctk.CTkLabel(
            self.root, text="تعداد فایل‌های نهایی: ۰",
            font=ctk.CTkFont(family=FONT_FAMILY, size=FONT_SIZE_NORMAL)
        )
        self.files_label.pack(pady=5)
        
        self.requests_label = ctk.CTkLabel(
            self.root, text="تعداد درخواست‌های تأیید شده: ۰",
            font=ctk.CTkFont(family=FONT_FAMILY, size=FONT_SIZE_NORMAL)
        )
        self.requests_label.pack(pady=5)
        
        self.downloads_label = ctk.CTkLabel(
            self.root, text="تعداد دانلودهای انجام شده: ۰",
            font=ctk.CTkFont(family=FONT_FAMILY, size=FONT_SIZE_NORMAL)
        )
        self.downloads_label.pack(pady=5)
        
        self.last_update_label = ctk.CTkLabel(
            self.root, text="آخرین بروزرسانی: --:--:--",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12)
        )
        self.last_update_label.pack(pady=10)
        
        # دکمه بروزرسانی دستی
        self.refresh_btn = ctk.CTkButton(
            self.root, text="بروزرسانی دستی", command=self.refresh_stats,
            font=ctk.CTkFont(family=FONT_FAMILY, size=FONT_SIZE_NORMAL),
            fg_color="#529e98", hover_color="#3d7a74"
        )
        self.refresh_btn.pack(pady=5)
        
        # دکمه اجرای برنامه اصلی (فقط زمانی فعال که سرویس متوقف باشد)
        self.launch_btn = ctk.CTkButton(
            self.root, text="🚀 اجرای  سرویس", command=self.launch_client_app,
            font=ctk.CTkFont(family=FONT_FAMILY, size=FONT_SIZE_NORMAL),
            fg_color="#2c7a4d", hover_color="#1e5a3a"
        )
        self.launch_btn.pack(pady=5)
        
        # دکمه توقف سرویس (فقط زمانی فعال که سرویس در حال اجرا باشد)
        self.stop_service_btn = ctk.CTkButton(
            self.root, text="🛑 توقف سرویس", command=self.stop_service,
            font=ctk.CTkFont(family=FONT_FAMILY, size=FONT_SIZE_NORMAL),
            fg_color="#d32f2f", hover_color="#b71c1c", state="disabled"
        )
        self.stop_service_btn.pack(pady=5)
        
        # نمایش مسیر برنامه اصلی
        if CLIENT_APP_PATH and os.path.exists(CLIENT_APP_PATH):
            path_label = ctk.CTkLabel(
                self.root, text=f"مسیر برنامه: {CLIENT_APP_PATH}",
                font=ctk.CTkFont(family=FONT_FAMILY, size=10), text_color="gray"
            )
            path_label.pack(pady=2)
        else:
            ctk.CTkLabel(
                self.root, text="⚠ مسیر برنامه اصلی در تنظیمات مشخص نشده یا فایل وجود ندارد",
                font=ctk.CTkFont(family=FONT_FAMILY, size=10), text_color="orange"
            ).pack(pady=2)
        
        # گزینه فعال/غیرفعال اعلان‌ها
        self.notif_var = ctk.BooleanVar(value=ENABLE_NOTIFICATIONS)
        self.notif_check = ctk.CTkCheckBox(
            self.root, text="فعال کردن اعلان‌ها", variable=self.notif_var,
            command=self.toggle_notifications,
            font=ctk.CTkFont(family=FONT_FAMILY, size=FONT_SIZE_NORMAL)
        )
        self.notif_check.pack(pady=5)
        
        # شروع چک کردن سرویس
        self.root.after(100, self.check_service_status)
        self.root.after(1000, self.refresh_stats_loop)
        
        # Tray
        self.create_tray_icon()
        self.tray_thread = threading.Thread(target=self.run_tray, daemon=True)
        self.tray_thread.start()
    
    def create_tray_icon(self):
        image = Image.new('RGB', (64, 64), color='green')
        draw = ImageDraw.Draw(image)
        draw.rectangle([16, 16, 48, 48], fill='white')
        self.tray_image = image
    
    def run_tray(self):
        menu = pystray.Menu(
            item("نمایش پنجره", self.show_window),
            item("🚀 اجرای برنامه", self.launch_client_app),
            item("🛑 توقف سرویس", self.stop_service),
            item("خروج", self.quit_app)
        )
        self.icon = pystray.Icon("service_monitor", self.tray_image, "مانیتور سرویس", menu=menu)
        self.icon.run()
    
    def show_window(self):
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
    
    def hide_window(self):
        self.root.withdraw()
    
    def quit_app(self):
        if self.icon:
            self.icon.stop()
        self.root.quit()
        sys.exit(0)
    
    def launch_client_app(self):
        if not CLIENT_APP_PATH:
            self.show_error("مسیر برنامه اصلی در فایل تنظیمات مشخص نشده است.")
            return
        if not os.path.exists(CLIENT_APP_PATH):
            self.show_error(f"فایل اجرایی یافت نشد:\n{CLIENT_APP_PATH}")
            return
        try:
            subprocess.Popen([CLIENT_APP_PATH], shell=True)
            self.show_info("برنامه اصلی با موفقیت اجرا شد.")
        except Exception as e:
            self.show_error(f"خطا در اجرای برنامه:\n{str(e)}")
    
    def stop_service(self):
        """پیدا کردن فرآیند سرویس (بر اساس نام یا پورت) و terminate کردن آن"""
        if not self.service_running:
            return
        try:
            # روش اول: پیدا کردن فرآیند از طریق PID ذخیره شده (اگر داشته باشیم)
            killed = False
            if self.service_pid:
                try:
                    proc = psutil.Process(self.service_pid)
                    proc.terminate()
                    proc.wait(timeout=5)
                    killed = True
                    self.send_notification("سرویس متوقف شد", "سرویس با موفقیت متوقف گردید.")
                except:
                    pass
            
            # روش دوم: جستجو بر اساس نام فرآیند
            if not killed:
                for proc in psutil.process_iter(['pid', 'name']):
                    try:
                        if proc.info['name'] and SERVICE_PROCESS_NAME.lower() in proc.info['name'].lower():
                            proc.terminate()
                            proc.wait(timeout=5)
                            killed = True
                            break
                    except:
                        continue
            
            # روش سوم: جستجو بر اساس پورت 5000
            if not killed:
                for conn in psutil.net_connections(kind='tcp'):
                    if conn.laddr.port == 5000 and conn.pid:
                        try:
                            proc = psutil.Process(conn.pid)
                            proc.terminate()
                            proc.wait(timeout=5)
                            killed = True
                            break
                        except:
                            continue
            
            if killed:
                # منتظر بمان تا وضعیت به روز شود
                time.sleep(2)
                self.check_service_status(manual=True)
                self.show_info("سرویس با موفقیت متوقف شد.")
            else:
                self.show_error("سرویس یافت نشد. ممکن است از قبل متوقف شده باشد.")
        except Exception as e:
            self.show_error(f"خطا در توقف سرویس:\n{str(e)}")
    
    def toggle_notifications(self):
        global ENABLE_NOTIFICATIONS
        ENABLE_NOTIFICATIONS = self.notif_var.get()
    
    def send_notification(self, title, message):
        if ENABLE_NOTIFICATIONS:
            try:
                notification.notify(
                    title=title,
                    message=message,
                    app_name="مانیتور سرویس",
                    timeout=5
                )
            except Exception as e:
                print(f"خطا در ارسال اعلان: {e}")
    
    def check_service_status(self, manual=False):
        """بررسی وضعیت سرویس از طریق health endpoint و نیز وجود فرآیند"""
        try:
            response = requests.get(f"{API_BASE_URL}/health", timeout=3)
            new_status = response.status_code == 200
            # اگر پاسخ 200 باشد، سرویس در حال اجراست
            if new_status:
                # پیدا کردن PID سرویس از طریق پورت
                try:
                    for conn in psutil.net_connections(kind='tcp'):
                        if conn.laddr.port == 5000 and conn.pid:
                            self.service_pid = conn.pid
                            break
                except:
                    pass
        except:
            new_status = False
            self.service_pid = None
        
        # تشخیص تغییر وضعیت برای ارسال اعلان
        if self.last_status is None:
            self.service_running = new_status
            self.last_status = new_status
            if new_status:
                self.send_notification("سرویس فعال شد", "سرویس تولید فایل نهایی در حال اجرا است.")
        else:
            if new_status != self.service_running:
                self.service_running = new_status
                if new_status:
                    self.send_notification("✅ سرویس شروع شد", "سرویس تولید فایل نهایی دوباره آنلاین شد.")
                else:
                    self.send_notification("❌ سرویس متوقف شد", "سرویس تولید فایل نهایی قطع شده است. لطفاً بررسی کنید.")
        
        # به‌روزرسانی UI
        if self.service_running:
            self.status_label.configure(text="وضعیت: ✅ در حال اجرا", text_color="green")
            self.launch_btn.configure(state="disabled")          # غیرفعال کردن دکمه اجرا
            self.stop_service_btn.configure(state="normal", fg_color="#d32f2f")  # فعال کردن دکمه توقف (قرمز)
        else:
            self.status_label.configure(text="وضعیت: ❌ قطع است", text_color="red")
            self.launch_btn.configure(state="normal")             # فعال کردن دکمه اجرا
            self.stop_service_btn.configure(state="disabled")     # غیرفعال کردن دکمه توقف
        
        self.last_status = self.service_running
        # اگر manual نبود، بعد از interval دوباره فراخوانی شود
        if not manual:
            self.root.after(REFRESH_INTERVAL * 1000, self.check_service_status)
    
    def refresh_stats(self):
        if not self.service_running:
            return
        try:
            response = requests.get(f"{API_BASE_URL}/stats", timeout=5)
            if response.status_code == 200:
                data = response.json()
                self.files_label.configure(text=f"تعداد فایل‌های نهایی: {data.get('total_files', 0):,}")
                self.requests_label.configure(text=f"تعداد درخواست‌های تأیید شده: {data.get('total_approved_requests', 0):,}")
                self.downloads_label.configure(text=f"تعداد دانلودهای انجام شده: {data.get('total_downloads', 0):,}")
                now = time.strftime("%H:%M:%S")
                self.last_update_label.configure(text=f"آخرین بروزرسانی: {now}")
        except Exception as e:
            self.files_label.configure(text="خطا در دریافت آمار")
    
    def refresh_stats_loop(self):
        self.refresh_stats()
        self.root.after(30000, self.refresh_stats_loop)
    
    def show_error(self, message):
        ctk.CTkMessagebox(self.root, title="خطا", message=message, icon="cancel")
    
    def show_info(self, message):
        ctk.CTkMessagebox(self.root, title="اطلاع", message=message, icon="info")
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    try:
        app = ServiceMonitorApp()
        app.run()
    except Exception as e:
        print(f"خطا در اجرای برنامه مانیتور: {e}")
        input("Enter را بزنید...")