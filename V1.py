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

# ================= 局域网配置区 =================
LAN_PWD_PATH = r"\\10.1.93.32\DT_HU_RDteam_F\视频\Z\密码\password.txt" 
LAN_LOG_PATH = r"\\10.1.93.32\DT_HU_RDteam_F\视频\Z\密码\log.txt"
LAN_UPDATE_SRC = r"\\10.1.93.32\DT_HU_RDteam_F\视频\Z\密码\update\摸鱼进站工具.exe"
# ===============================================

BARCODE_HISTORY = set()
SCAN_BUFFER = []
LAST_KEY_TIME = 0
SCAN_SPEED_THRESHOLD = 0.05 
kb_controller = Controller()

class ClassicV4:
    def __init__(self, root):
        self.root = root
        self.root.title("摸鱼进站 v4.0 (稳定版)")
        self.root.geometry("400x750")
        self.root.attributes("-topmost", True)
        self.root.configure(bg="#f5f5f5")

        # 1. 列表区 (Treeview)
        tree_frame = tk.Frame(self.root)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.tree = ttk.Treeview(tree_frame, columns=("check", "sn"), show="headings", height=15)
        self.tree.heading("check", text="选")
        self.tree.heading("sn", text="序列号 SN (18位)")
        self.tree.column("check", width=40, anchor="center")
        self.tree.column("sn", width=330)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        sb = tk.Scrollbar(tree_frame, command=self.tree.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.config(yscrollcommand=sb.set)
        self.tree.bind("<Button-1>", self.toggle_check)

        # 2. 按钮操作区
        btn_f = tk.Frame(self.root)
        btn_f.pack(fill=tk.X, padx=10)
        tk.Button(btn_f, text="📋 粘贴排序", command=self.paste_sn, bg="#E1F5FE").pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        tk.Button(btn_f, text="❌ 删除勾选", command=self.delete_checked, bg="#FFEBEE").pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        tk.Button(btn_f, text="🗑️ 清空", command=lambda: self.tree.delete(*self.tree.get_children())).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

        # 3. 速度控制滑块
        speed_frame = tk.LabelFrame(self.root, text="🚀 录入参数调节", pady=5)
        speed_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.s_interval = tk.Scale(speed_frame, from_=0.1, to=2.0, resolution=0.1, orient=tk.HORIZONTAL, label="每条间隔(秒)")
        self.s_interval.set(0.7)
        self.s_interval.pack(fill=tk.X, padx=10)

        # 拦截回跳开关
        self.enable_pullback = tk.BooleanVar(value=True)
        tk.Checkbutton(self.root, text="重复拦截拉回上一格 (Shift+Tab + Ctrl+A)", 
                       variable=self.enable_pullback, fg="red", font=("微软雅黑", 9, "bold")).pack(pady=2)

        # 4. 实时日志与公告区
        log_f = tk.LabelFrame(self.root, text="📢 状态与公告", padx=5, pady=5)
        log_f.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.log_display = scrolledtext.ScrolledText(log_f, height=6, font=("Consolas", 9), bg="#1e1e1e", fg="white")
        self.log_display.pack(fill=tk.BOTH, expand=True)

        # 5. 开始按钮
        tk.Button(self.root, text="🔥 开始自动化录入 (5s准备)", bg="#2E7D32", fg="white", 
                  font=("微软雅黑", 10, "bold"), pady=10, command=self.start_work).pack(fill=tk.X, padx=10, pady=10)

        self.refresh_lan_log()

    def toggle_check(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            cur = self.tree.set(item, "check")
            self.tree.set(item, "check", "☑" if cur == "☐" else "☐")

    def paste_sn(self):
        try:
            data = self.root.clipboard_get()
            sns = list(set([l.strip() for l in data.split('\n') if l.strip()]))
            sns.sort()
            for s in sns:
                self.tree.insert("", tk.END, values=("☐", s))
        except: pass

    def delete_checked(self):
        for i in self.tree.get_children():
            if self.tree.set(i, "check") == "☑":
                self.tree.delete(i)

    def refresh_lan_log(self):
        def read():
            try:
                if os.path.exists(LAN_LOG_PATH):
                    with open(LAN_LOG_PATH, "r", encoding="utf-8-sig") as f:
                        content = f.read()
                    self.root.after(0, lambda: self.log_display.insert(tk.END, f"\n---内网公告---\n{content}\n"))
            except: pass
        threading.Thread(target=read, daemon=True).start()

    def start_work(self):
        items = self.tree.get_children()
        if not items: return
        self.root.attributes("-alpha", 0.3)
        threading.Thread(target=self.work_logic, args=(items,), daemon=True).start()

    def work_logic(self, items):
        time.sleep(5)
        interval = self.s_interval.get()
        for i in items:
            sn = self.tree.set(i, "sn")
            # 执行录入
            kb_controller.press(Key.ctrl)
            kb_controller.press('a')
            kb_controller.release('a')
            time.sleep(0.1)
            self.root.after(0, lambda x=sn: [self.root.clipboard_clear(), self.root.clipboard_append(x)])
            time.sleep(0.1)
            kb_controller.press('v')
            kb_controller.release('v')
            kb_controller.release(Key.ctrl)
            time.sleep(0.2)
            kb_controller.press(Key.enter)
            kb_controller.release(Key.enter)
            time.sleep(interval)
        self.root.after(0, lambda: [self.root.attributes("-alpha", 1.0), winsound.Beep(1000, 300)])

    def update_monitor(self, code, is_dup):
        ts = time.strftime("%H:%M:%S")
        msg = f"[{ts}] {'[DUP]' if is_dup else '[OK]'} {code}\n"
        self.log_display.insert(tk.END, msg)
        self.log_display.see(tk.END)
        if is_dup:
            winsound.Beep(1500, 600)
            if self.enable_pullback.get():
                # 修复语法错误：拆分 with 语句
                with kb_controller.pressed(Key.shift):
                    kb_controller.press(Key.tab)
                    kb_controller.release(Key.tab)
                
                time.sleep(0.15)
                
                with kb_controller.pressed(Key.ctrl):
                    kb_controller.press('a')
                    kb_controller.release('a')

# --- 验证与更新系统 ---
def get_file_md5(f):
    if not os.path.exists(f): return None
    h = hashlib.md5()
    with open(f, "rb") as _f:
        for c in iter(lambda: _f.read(4096), b""): h.update(c)
    return h.hexdigest()

def check_login():
    lw = tk.Tk(); lw.title("验证"); lw.geometry("240x120")
    lw.eval('tk::PlaceWindow . center')
    tk.Label(lw, text="授权码:").pack(pady=5)
    pw_ent = tk.Entry(lw, show="*"); pw_ent.pack(); pw_ent.focus_set()
    def go():
        try:
            with open(LAN_PWD_PATH, "r", encoding="utf-8-sig") as f:
                if pw_ent.get() == f.read().strip():
                    lw.withdraw()
                    # 检查更新
                    src, cur = LAN_UPDATE_SRC, (sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__))
                    if os.path.exists(src) and get_file_md5(src) != get_file_md5(cur):
                        if messagebox.askyesno("更新", "发现新版本，是否升级？"):
                            with open("updater.bat", "w") as f:
                                f.write(f'@echo off\ntimeout /t 1\ncopy /y "{src}" "{cur}"\nstart "" "{cur}"\ndel %0')
                            subprocess.Popen("updater.bat", shell=True); sys.exit()
                    lw.destroy(); start_app()
                else: messagebox.showerror("!", "码错")
        except: messagebox.showerror("!", "内网断了")
    tk.Button(lw, text="进入", command=go).pack(pady=10)
    lw.bind('<Return>', lambda e: go()); lw.mainloop()

def start_app():
    global app
    root = tk.Tk()
    app = ClassicV4(root)
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
                if 'app' in globals():
                    app.root.after(0, lambda: app.update_monitor(barcode, is_dup))
            SCAN_BUFFER = []
        elif hasattr(key, 'char') and key.char:
            if interval > SCAN_SPEED_THRESHOLD: SCAN_BUFFER = []
            SCAN_BUFFER.append(key.char)
    except: pass

if __name__ == "__main__":
    check_login()
