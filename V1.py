import time
import threading
import winsound
import tkinter as tk
from tkinter import ttk, scrolledtext
from pynput import keyboard
from pynput.keyboard import Controller, Key

# --- 全局变量 ---
BARCODE_HISTORY = set()
SCAN_BUFFER = []
LAST_KEY_TIME = 0
SCAN_SPEED_THRESHOLD = 0.05 
kb_controller = Controller()

class ProfessionalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("摸鱼全能版 v3.5")
        self.root.geometry("380x680")
        self.root.attributes("-topmost", True)
        
        # --- 1. 自动录入区 ---
        entry_f = tk.LabelFrame(self.root, text=" ⚡ 批量进站自动化 ", font=("微软雅黑", 9, "bold"), pady=5)
        entry_f.pack(fill=tk.X, padx=10, pady=5)
        
        btn_f = tk.Frame(entry_f)
        btn_f.pack(fill=tk.X, padx=5)
        tk.Button(btn_f, text="📋 粘贴SN", command=self.paste_sn, bg="#e3f2fd").pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        tk.Button(btn_f, text="🗑️ 清空", command=lambda: self.sn_list.delete(0, tk.END)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        
        self.sn_list = tk.Listbox(entry_f, height=5, font=("Consolas", 9))
        self.sn_list.pack(fill=tk.X, padx=5, pady=5)
        
        tk.Button(entry_frame:=entry_f, text="🔥 开始执行 (5秒准备)", bg="#2e7d32", fg="white", 
                  font=("微软雅黑", 9, "bold"), command=self.start_entry_thread).pack(fill=tk.X, padx=5)

        # --- 2. 扫码监控区 ---
        mon_f = tk.LabelFrame(self.root, text=" 🛡️ 扫码防重监控 ", font=("微软雅黑", 9, "bold"), pady=5)
        mon_f.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 【新增：拦截拉回开关】
        ctrl_f = tk.Frame(mon_f)
        ctrl_f.pack(fill=tk.X, padx=5)
        self.enable_pullback = tk.BooleanVar(value=True) # 默认开启
        tk.Checkbutton(ctrl_f, text="开启重复拉回 (Shift+Tab & Ctrl+A)", 
                       variable=self.enable_pullback, fg="#d32f2f", font=("微软雅黑", 9, "bold")).pack(side=tk.LEFT)

        self.status_bar = tk.Label(mon_f, text="🟢 监控中...", bg="#4caf50", fg="white", font=("微软雅黑", 10))
        self.status_bar.pack(fill=tk.X, padx=5, pady=5)

        self.log_area = scrolledtext.ScrolledText(mon_f, height=12, font=("Consolas", 8))
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # --- 3. 底部工具栏 ---
        bottom_f = tk.Frame(self.root)
        bottom_f.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(bottom_f, text="置顶:", font=("微软雅黑", 8)).pack(side=tk.LEFT)
        self.stay_top = tk.BooleanVar(value=True)
        tk.Checkbutton(bottom_f, variable=self.stay_top, 
                       command=lambda: self.root.attributes("-topmost", self.stay_top.get())).pack(side=tk.LEFT)
        tk.Button(bottom_f, text="🗑️ 清空记录", command=self.clear_logs, font=("微软雅黑", 8)).pack(side=tk.RIGHT)

    def paste_sn(self):
        try:
            for s in self.root.clipboard_get().split('\n'):
                if s.strip(): self.sn_list.insert(tk.END, s.strip())
        except: pass

    def clear_logs(self): self.log_area.delete('1.0', tk.END); BARCODE_HISTORY.clear()

    def start_entry_thread(self):
        sns = self.sn_list.get(0, tk.END)
        if sns:
            self.root.attributes("-alpha", 0.4)
            threading.Thread(target=self._run_entry, args=(sns,), daemon=True).start()

    def _run_entry(self, sns):
        time.sleep(5)
        for sn in sns:
            kb_controller.press(Key.ctrl); kb_controller.press('a'); kb_controller.release('a')
            time.sleep(0.1)
            self.root.after(0, lambda x=sn: [self.root.clipboard_clear(), self.root.clipboard_append(x)])
            time.sleep(0.1)
            kb_controller.press('v'); kb_controller.release('v'); kb_controller.release(Key.ctrl)
            time.sleep(0.2)
            kb_controller.press(Key.enter); kb_controller.release(Key.enter)
            time.sleep(0.8)
        self.root.after(0, lambda: [self.root.attributes("-alpha", 1.0), winsound.Beep(1000, 300)])

    def update_monitor(self, code, is_dup):
        ts = time.strftime("%H:%M:%S")
        if is_dup:
            self.status_bar.config(text=f"❌ 重复拦截: {code}", bg="#f44336")
            self.log_area.insert(tk.END, f"[{ts}] 重复: {code}\n", "err")
            self.log_area.tag_config("err", foreground="red")
            winsound.Beep(1500, 600)
            # 执行拉回逻辑 (判断开关)
            if self.enable_pullback.get():
                self.execute_pullback()
        else:
            self.status_bar.config(text=f"✅ 扫描通过: {code}", bg="#4caf50")
            self.log_area.insert(tk.END, f"[{ts}] 通过: {code}\n")
        self.log_area.see(tk.END)

    def execute_pullback(self):
        """执行B方法：拉回并全选"""
        with kb_controller.pressed(Key.shift):
            kb_controller.press(Key.tab); kb_controller.release(Key.tab)
        time.sleep(0.15) # 稍微增加延迟提高稳定性
        with kb_controller.pressed(Key.ctrl):
            kb_controller.press('a'); kb_controller.release('a')

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
                app.root.after(0, app.update_monitor, barcode, is_dup)
            SCAN_BUFFER = []
        elif hasattr(key, 'char') and key.char:
            if interval > SCAN_SPEED_THRESHOLD: SCAN_BUFFER = []
            SCAN_BUFFER.append(key.char)
    except: pass

if __name__ == "__main__":
    root = tk.Tk()
    style = ttk.Style(root); style.theme_use('vista')
    app = ProfessionalApp(root)
    threading.Thread(target=lambda: keyboard.Listener(on_press=on_press).start(), daemon=True).start()
    root.mainloop()
