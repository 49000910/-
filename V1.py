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

class FinalProApp:
    def __init__(self, root):
        self.root = root
        self.root.title("摸鱼工具箱 v4.2")
        self.root.geometry("400x750")
        self.root.attributes("-topmost", True, "-alpha", 0.9)
        self.root.configure(bg="#121212")
        
        self.dark_idle, self.dark_green, self.dark_red = "#1e1e1e", "#004d00", "#660000"
        self.text_fg = "#ffffff"

        # --- 1. 批量录入区 ---
        entry_f = tk.LabelFrame(self.root, text=" ⚡ 录入清单 ", font=("微软雅黑", 8), bg=self.dark_idle, fg="#888")
        entry_f.pack(fill=tk.BOTH, expand=True, padx=8, pady=2)
        
        btn_f = tk.Frame(entry_f, bg=self.dark_idle)
        btn_f.pack(fill=tk.X, padx=2, pady=1)
        tk.Button(btn_f, text="📋 粘贴", command=self.paste_sn, bg="#333", fg="white", bd=0, font=("微软雅黑", 8)).pack(side=tk.RIGHT, padx=2)
        tk.Button(btn_f, text="🗑️ 清空", command=lambda: self.sn_list.delete(0, tk.END), bg="#333", fg="white", bd=0, font=("微软雅黑", 8)).pack(side=tk.RIGHT, padx=2)
        
        list_f = tk.Frame(entry_f, bg=self.dark_idle)
        list_f.pack(fill=tk.BOTH, expand=True, padx=2)
        self.sb1 = tk.Scrollbar(list_f)
        self.sb1.pack(side=tk.RIGHT, fill=tk.Y)
        self.sn_list = tk.Listbox(list_f, bg="#121212", fg=self.text_fg, bd=0, font=("Consolas", 9), yscrollcommand=self.sb1.set)
        self.sn_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.sb1.config(command=self.sn_list.yview)

        # 功能控制行
        ctrl_p = tk.Frame(entry_f, bg=self.dark_idle)
        ctrl_p.pack(fill=tk.X, padx=2, pady=2)
        tk.Button(ctrl_p, text="❌ 删除选中", bg="#421010", fg="#ff9999", bd=0, font=("微软雅黑", 8), command=self.delete_selected).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=1)
        tk.Button(ctrl_p, text="🔥 开始录入 (5s)", bg="#1b5e20", fg="white", bd=0, font=("微软雅黑", 8, "bold"), command=self.start_entry_thread).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=1)
        
        # 精简参数调节区 (限制范围 0.1-1.0s)
        settings_f = tk.Frame(entry_f, bg=self.dark_idle)
        settings_f.pack(fill=tk.X, padx=2, pady=2)
        
        self.enable_pullback = tk.BooleanVar(value=True)
        tk.Checkbutton(settings_f, text="拦截回跳", variable=self.enable_pullback, bg=self.dark_idle, fg="#ffab00", font=("微软雅黑", 8), selectcolor="#000").pack(side=tk.LEFT)
        
        def create_s(label, default):
            f = tk.Frame(settings_f, bg=self.dark_idle)
            f.pack(side=tk.RIGHT, padx=2)
            tk.Label(f, text=label, bg=self.dark_idle, fg="#666", font=("微软雅黑", 7)).pack(side=tk.LEFT)
            s = tk.Scale(f, from_=0.1, to=1.0, resolution=0.1, orient=tk.HORIZONTAL, bg=self.dark_idle, fg="#aaa", bd=0, highlightthickness=0, length=55, font=("Arial", 7), showvalue=True)
            s.set(default); s.pack(side=tk.LEFT); return s

        self.s_double_enter = create_s("双回车:", 0.5)
        self.s_enter_speed = create_s("回车时间:", 0.8)

        # --- 2. 扫码监控区 ---
        mon_f = tk.LabelFrame(self.root, text=" 🛡️ 扫描防重监控 ", font=("微软雅黑", 8), bg=self.dark_idle, fg="#888")
        mon_f.pack(fill=tk.X, padx=8, pady=2)
        self.status_bar = tk.Label(mon_f, text="等待扫描...", bg="#222", fg="white", font=("微软雅黑", 9, "bold"))
        self.status_bar.pack(fill=tk.X, padx=2, pady=1)
        self.log_area = scrolledtext.ScrolledText(mon_f, height=8, bg="#222", fg="white", font=("Consolas", 8), bd=0)
        self.log_area.pack(fill=tk.X, padx=2, pady=2)

        # --- 3. 内网公告 ---
        lan_f = tk.LabelFrame(self.root, text=" 📢 内网公告 ", font=("微软雅黑", 7), bg="#000", fg="#00b0ff")
        lan_f.pack(fill=tk.X, padx=8, pady=2)
        self.lan_display = tk.Text(lan_f, height=3, bg="#000", fg="#00b0ff", font=("微软雅黑", 7), bd=0, padx=5)
        self.lan_display.pack(fill=tk.X)
        self.refresh_lan_log()

        # --- 4. 底部栏 ---
        bottom_f = tk.Frame(self.root, bg="#000")
        bottom_f.pack(fill=tk.X, side=tk.BOTTOM)
        self.stay_top = tk.BooleanVar(value=True)
        tk.Checkbutton(bottom_f, text="始终置顶", variable=self.stay_top, bg="#000", fg="#444", font=("微软雅黑", 7), command=lambda: self.root.attributes("-topmost", self.stay_top.get())).pack(side=tk.LEFT, padx=5)
        tk.Button(bottom_f, text="清空记录", command=self.clear_logs, font=("微软雅黑", 7), bd=0, bg="#000", fg="#444").pack(side=tk.RIGHT, padx=5)

    def refresh_lan_log(self):
        def read():
            try:
                if os.path.exists(LAN_LOG_PATH):
                    with open(LAN_LOG_PATH, "r", encoding="utf-8-sig") as f:
                        content = f.read()
                    self.root.after(0, lambda: self._update_lan_ui(content))
            except: pass
        threading.Thread(target=read, daemon=True).start()

    def _update_lan_ui(self, txt):
        self.lan_display.config(state=tk.NORMAL)
        self.lan_display.delete('1.0', tk.END); self.lan_display.insert(tk.END, txt)
        self.lan_display.config(state=tk.DISABLED)

    def delete_selected(self):
        idx = self.sn_list.curselection()
        if idx: self.sn_list.delete(idx)

    def paste_sn(self):
        try:
            for s in self.root.clipboard_get().split('\n'):
                if s.strip(): self.sn_list.insert(tk.END, s.strip())
        except: pass

    def clear_logs(self):
        self.log_area.config(bg="#222"); self.status_bar.config(bg="#222", text="等待扫描...")
        self.log_area.delete('1.0', tk.END); BARCODE_HISTORY.clear()

    def start_entry_thread(self):
        sns = self.sn_list.get(0, tk.END)
        if sns:
            self.root.attributes("-alpha", 0.3)
            threading.Thread(target=self._run_entry, args=(sns,), daemon=True).start()

    def _run_entry(self, sns):
        time.sleep(5)
        for sn in sns:
            kb_controller.press(Key.ctrl); kb_controller.press('a'); kb_controller.release('a'); time.sleep(0.1)
            self.root.after(0, lambda x=sn: [self.root.clipboard_clear(), self.root.clipboard_append(x)])
            time.sleep(0.1); kb_controller.press('v'); kb_controller.release('v'); kb_controller.release(Key.ctrl)
            time.sleep(0.2); kb_controller.press(Key.enter); kb_controller.release(Key.enter)
            time.sleep(self.s_double_enter.get())
            kb_controller.press(Key.enter); kb_controller.release(Key.enter)
            time.sleep(self.s_enter_speed.get())
        self.root.after(0, lambda: [self.root.attributes("-alpha", 0.9), winsound.Beep(1000, 300)])

    def update_monitor(self, code, is_dup):
        ts = time.strftime("%H:%M:%S")
        c = self.dark_red if is_dup else self.dark_green
        self.status_bar.config(text=f"{'!! 重复' if is_dup else 'OK扫描'}: {code}", bg=c)
        self.log_area.config(bg=c)
        self.log_area.insert(tk.END, f"[{ts}] {'DUP' if is_dup else 'PASS'} -> {code}\n")
        self.log_area.see(tk.END)
        if is_dup:
            winsound.Beep(1500, 600)
            if self.enable_pullback.get():
                with kb_controller.pressed(Key.shift): kb_controller.press(Key.tab); kb_controller.release(Key.tab)
                time.sleep(0.15); with kb_controller.pressed(Key.ctrl): kb_controller.press('a'); kb_controller.release('a')

def get_file_md5(f):
    if not os.path.exists(f): return None
    h = hashlib.md5()
    with open(f, "rb") as _f:
        for c in iter(lambda: _f.read(4096), b""): h.update(c)
    return h.hexdigest()

def check_update_and_login():
    login_w = tk.Tk(); login_w.title("验证"); login_w.geometry("240x120")
    login_w.eval('tk::PlaceWindow . center')
    tk.Label(login_w, text="请输入授权码:").pack(pady=5)
    pw_ent = tk.Entry(login_w, show="*"); pw_ent.pack(); pw_ent.focus_set()

    def do_login():
        try:
            with open(LAN_PWD_PATH, "r", encoding="utf-8-sig") as f:
                if pw_ent.get() == f.read().strip():
                    login_w.withdraw()
                    src = LAN_UPDATE_SRC
                    cur = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
                    if os.path.exists(src) and get_file_md5(src) != get_file_md5(cur):
                        if messagebox.askyesno("更新", "检测到新版本，是否升级？"):
                            with open("updater.bat", "w") as f:
                                f.write(f'@echo off\ntimeout /t 1\ncopy /y "{src}" "{cur}"\nstart "" "{cur}"\ndel %0')
                            subprocess.Popen("updater.bat", shell=True); sys.exit()
                    login_w.destroy(); start_main_app()
                else: messagebox.showerror("!", "授权码错误")
        except: messagebox.showerror("!", "无法连接内网服务器")

    tk.Button(login_w, text="登录", command=do_login, width=10).pack(pady=10)
    login_w.bind('<Return>', lambda e: do_login()); login_w.mainloop()

def start_main_app():
    global app
    root = tk.Tk()
    app = FinalProApp(root)
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
                if not is_dup: BARCODE_HISTORY.add(barcode)
                app.root.after(0, lambda: app.update_monitor(barcode, is_dup))
            SCAN_BUFFER = []
        elif hasattr(key, 'char') and key.char:
            if interval > SCAN_SPEED_THRESHOLD: SCAN_BUFFER = []
            SCAN_BUFFER.append(key.char)
    except: pass

if __name__ == "__main__":
    check_update_and_login()
