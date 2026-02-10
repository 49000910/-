import ctypes
import time
import os
import threading
import winsound
import hashlib
import sys
import subprocess
import tkinter as tk
from tkinter import messagebox, ttk, scrolledtext
from pynput import keyboard
from pynput.keyboard import Controller, Key

# ================= 局域网配置 =================
LAN_PWD_PATH = r"\\10.1.93.32\DT_HU_RDteam_F\视频\Z\密码\password.txt" 
LAN_LOG_PATH = r"\\10.1.93.32\DT_HU_RDteam_F\视频\Z\密码\log.txt"
LAN_UPDATE_SRC = r"\\10.1.93.32\DT_HU_RDteam_F\视频\Z\密码\update\摸鱼工具箱.exe"
# =============================================

BARCODE_HISTORY = set()
SCAN_BUFFER = []
LAST_KEY_TIME = 0
SCAN_SPEED_THRESHOLD = 0.05 
kb_controller = Controller()

class MainMonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("工具箱监控 v4.6")
        self.root.geometry("400x650") 
        self.root.attributes("-topmost", True, "-alpha", 0.95)
        self.root.configure(bg="#121212")
        
        self.dark_idle = "#1e1e1e"
        self.dark_green = "#004d00"
        self.dark_red = "#660000"
        
        # --- 1. 顶部状态 ---
        self.status_bar = tk.Label(self.root, text="🛡️ 实时防重监控中", bg="#222", 
                                   fg="white", font=("微软雅黑", 10, "bold"), pady=8)
        self.status_bar.pack(fill=tk.X)

        # --- 2. 实时日志 (18位对齐) ---
        self.log_area = scrolledtext.ScrolledText(self.root, height=15, bg="#222", 
                                                  fg="white", font=("Consolas", 10), bd=0)
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # --- 3. 控制面板 ---
        ctrl_f = tk.Frame(self.root, bg=self.dark_idle, pady=5)
        ctrl_f.pack(fill=tk.X, padx=10)
        
        # 弹出清单按钮
        tk.Button(ctrl_f, text="📋 打开批量录入清单", command=self.open_entry_window, 
                  bg="#1b5e20", fg="white", font=("微软雅黑", 9, "bold"), bd=0, height=2).pack(fill=tk.X, pady=5)
        
        settings_f = tk.Frame(ctrl_f, bg=self.dark_idle)
        settings_f.pack(fill=tk.X)
        self.enable_pullback = tk.BooleanVar(value=True)
        tk.Checkbutton(settings_f, text="重复拦截拉回", variable=self.enable_pullback, 
                       bg=self.dark_idle, fg="#ffab00", font=("微软雅黑", 8), selectcolor="#000").pack(side=tk.LEFT)
        
        self.s_enter_speed = tk.Scale(settings_f, from_=0.1, to=1.0, resolution=0.1, 
                                      orient=tk.HORIZONTAL, bg=self.dark_idle, fg="#aaa", 
                                      bd=0, highlightthickness=0, length=100, font=("Arial", 7), label="执行速度")
        self.s_enter_speed.set(0.8)
        self.s_enter_speed.pack(side=tk.RIGHT)

        # --- 4. 内网通知 ---
        lan_f = tk.LabelFrame(self.root, text=" 📢 内网公告 ", font=("微软雅黑", 7), bg="#000", fg="#00b0ff")
        lan_f.pack(fill=tk.X, padx=10, pady=5)
        self.lan_display = tk.Text(lan_f, height=3, bg="#000", fg="#00b0ff", font=("微软雅黑", 7), bd=0)
        self.lan_display.pack(fill=tk.X)
        self.refresh_lan_log()

    def open_entry_window(self):
        """弹出优化后的 18位 SN 录入窗口"""
        top = tk.Toplevel(self.root)
        top.title("录入清单")
        top.geometry("420x550") 
        top.attributes("-topmost", True)
        top.configure(bg="#1e1e1e")

        list_f = tk.Frame(top, bg="#1e1e1e")
        list_f.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        sb = tk.Scrollbar(list_f)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 18位SN专用字体 Consolas
        sn_list = tk.Listbox(list_f, bg="#121212", fg="#ddd", font=("Consolas", 11), 
                             bd=0, yscrollcommand=sb.set, selectbackground="#333")
        sn_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.config(command=sn_list.yview)

        btn_f = tk.Frame(top, bg="#1e1e1e")
        btn_f.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Button(btn_f, text="📋 粘贴导入", command=lambda: self.paste_to_list(sn_list), 
                  bg="#333", fg="white", width=12).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_f, text="❌ 选中排除", command=lambda: self.delete_from_list(sn_list), 
                  bg="#421010", fg="#ff9999", width=12).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_f, text="🔥 开始执行", command=lambda: self.start_auto_entry(sn_list, top), 
                  bg="#1b5e20", fg="white", font=("微软雅黑", 9, "bold")).pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=5)

    def paste_to_list(self, lb):
        try:
            data = self.root.clipboard_get()
            for line in data.split('\n'):
                if line.strip():
                    lb.insert(tk.END, line.strip())
        except:
            pass

    def delete_from_list(self, lb):
        """选中不录入逻辑：直接从当前清空排除"""
        indices = lb.curselection()
        for i in reversed(indices):
            lb.delete(i)

    def start_auto_entry(self, lb, window):
        sns = lb.get(0, tk.END)
        if not sns:
            return
        window.withdraw()
        self.root.attributes("-alpha", 0.3)
        threading.Thread(target=self._run_logic, args=(sns,), daemon=True).start()

    def _run_logic(self, sns):
        time.sleep(5)
        interval = self.s_enter_speed.get()
        for sn in sns:
            # 1. 全选
            kb_controller.press(Key.ctrl)
            kb_controller.press('a')
            kb_controller.release('a')
            time.sleep(0.1)
            
            # 2. 写入剪贴板 (异步安全)
            self.root.after(0, lambda x=sn: [self.root.clipboard_clear(), self.root.clipboard_append(x)])
            time.sleep(0.1)
            
            # 3. 粘贴并提交
            kb_controller.press('v')
            kb_controller.release('v')
            kb_controller.release(Key.ctrl)
            time.sleep(0.2)
            
            kb_controller.press(Key.enter)
            kb_controller.release(Key.enter)
            time.sleep(interval)
            
        self.root.after(0, lambda: [self.root.attributes("-alpha", 0.95), winsound.Beep(1000, 300)])
        messagebox.showinfo("完成", "任务结束")

    def refresh_lan_log(self):
        def read():
            try:
                if os.path.exists(LAN_LOG_PATH):
                    with open(LAN_LOG_PATH, "r", encoding="utf-8-sig") as f:
                        content = f.read()
                    self.root.after(0, lambda: self._update_lan_ui(content))
            except:
                pass
        threading.Thread(target=read, daemon=True).start()

    def _update_lan_ui(self, txt):
        self.lan_display.config(state=tk.NORMAL)
        self.lan_display.delete('1.0', tk.END)
        self.lan_display.insert(tk.END, txt)
        self.lan_display.config(state=tk.DISABLED)

    def clear_logs(self):
        self.log_area.config(bg="#222")
        self.status_bar.config(bg="#222", text="等待扫描...")
        self.log_area.delete('1.0', tk.END)
        BARCODE_HISTORY.clear()

    def update_monitor(self, code, is_dup):
        ts = time.strftime("%H:%M:%S")
        c = self.dark_red if is_dup else self.dark_green
        self.status_bar.config(text=f"{'!! 重复' if is_dup else 'OK 扫描'}: {code}", bg=c)
        self.log_area.config(bg=c)
        self.log_area.insert(tk.END, f"[{ts}] {'DUP' if is_dup else 'PASS'} -> {code}\n")
        self.log_area.see(tk.END)
        if is_dup:
            winsound.Beep(1500, 600)
            if self.enable_pullback.get():
                # 拉回逻辑：语法拆分确保 Git 打包成功
                with kb_controller.pressed(Key.shift):
                    kb_controller.press(Key.tab)
                    kb_controller.release(Key.tab)
                
                time.sleep(0.15)
                
                with kb_controller.pressed(Key.ctrl):
                    kb_controller.press('a')
                    kb_controller.release('a')

# --- 验证更新逻辑 ---
def get_file_md5(f):
    if not os.path.exists(f): return None
    h = hashlib.md5()
    with open(f, "rb") as _f:
        for c in iter(lambda: _f.read(4096), b""): h.update(c)
    return h.hexdigest()

def check_update_and_login():
    lw = tk.Tk()
    lw.title("验证")
    lw.geometry("240x120")
    lw.eval('tk::PlaceWindow . center')
    tk.Label(lw, text="请输入授权码:").pack(pady=5)
    pw_ent = tk.Entry(lw, show="*")
    pw_ent.pack()
    pw_ent.focus_set()

    def do_login():
        try:
            with open(LAN_PWD_PATH, "r", encoding="utf-8-sig") as f:
                if pw_ent.get() == f.read().strip():
                    lw.withdraw()
                    # 检查更新
                    src = LAN_UPDATE_SRC
                    cur = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
                    if os.path.exists(src) and get_file_md5(src) != get_file_md5(cur):
                        if messagebox.askyesno("更新", "发现新版本，是否升级？"):
                            with open("updater.bat", "w") as f:
                                f.write(f'@echo off\ntimeout /t 1\ncopy /y "{src}" "{cur}"\nstart "" "{cur}"\ndel %0')
                            subprocess.Popen("updater.bat", shell=True)
                            sys.exit()
                    lw.destroy()
                    start_main()
                else:
                    messagebox.showerror("!", "授权码错误")
        except:
            messagebox.showerror("!", "内网服务器连接失败")

    tk.Button(lw, text="登录系统", command=do_login, width=10).pack(pady=10)
    lw.bind('<Return>', lambda e: do_login())
    lw.mainloop()

def start_main():
    global app
    root = tk.Tk()
    app = MainMonitorApp(root)
    threading.Thread(target=lambda: keyboard.Listener(on_press=on_press).start(), daemon=True).start()
    root.mainloop()

def on_press(key):
    global LAST_KEY_TIME, SCAN_BUFFER
    now = time.time()
    interval = now - LAST_KEY_TIME
    LAST_KEY_TIME = now
    try:
        if key == Key.enter:
            barcode = "".join(SCAN_BUFFER).strip()
            if barcode:
                is_dup = barcode in BARCODE_HISTORY
                if not is_dup:
                    BARCODE_HISTORY.add(barcode)
                if 'app' in globals():
                    app.root.after(0, lambda: app.update_monitor(barcode, is_dup))
            SCAN_BUFFER = []
        elif hasattr(key, 'char') and key.char:
            if interval > SCAN_SPEED_THRESHOLD:
                SCAN_BUFFER = []
            SCAN_BUFFER.append(key.char)
    except:
        pass

if __name__ == "__main__":
    check_update_and_login()
