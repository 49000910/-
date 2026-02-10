import time
import threading
import winsound
import os
import sys
import tkinter as tk
from tkinter import messagebox
from pynput import keyboard
from pynput.keyboard import Controller, Key

# --- 核心配置 ---
HISTORY_FILE = "barcode_history.txt"
BARCODE_HISTORY = set()
SCAN_BUFFER = []
LAST_KEY_TIME = 0
SCAN_SPEED_THRESHOLD = 0.05 
kb_controller = Controller()

if os.path.exists(HISTORY_FILE):
    with open(HISTORY_FILE, "r", encoding="utf-8") as f:
        BARCODE_HISTORY = set(line.strip() for line in f if line.strip())

class UltimateThinMonitor:
    def __init__(self, root):
        self.root = root
        self.root.title("智能拉回防重助手 v4.0")
        self.root.geometry("380x300") 
        self.root.attributes("-topmost", True) # 强制锁定悬浮置顶
        self.root.attributes("-alpha", 0.92) 
        self.root.overrideredirect(True)

        # 1. 软件名标题栏
        self.title_bar = tk.Frame(self.root, bg="#2c3e50", height=22)
        self.title_bar.pack(fill=tk.X)
        tk.Label(self.title_bar, text=" 🛡️ 智能拉回防重助手 v4.0", font=("微软雅黑", 8, "bold"), fg="white", bg="#2c3e50").pack(side=tk.LEFT)
        tk.Button(self.title_bar, text="×", bg="#2c3e50", fg="white", bd=0, command=self.safe_exit, font=("Arial", 9)).pack(side=tk.RIGHT, padx=5)
        tk.Button(self.title_bar, text="—", bg="#2c3e50", fg="white", bd=0, command=self.minimize, font=("Arial", 9)).pack(side=tk.RIGHT, padx=5)

        # 2. 核心交互区
        self.main_f = tk.Frame(self.root, bg="#90ee90", pady=2) 
        self.main_f.pack(fill=tk.BOTH, expand=True)

        # 三参数横排调节区
        params_f = tk.Frame(self.main_f, bg="#90ee90")
        params_f.pack(fill=tk.X, padx=8, pady=5)
        spin_opt = {"font": ("Consolas", 9), "width": 3, "from_": 0.0, "to": 5.0, "increment": 0.1, "bd": 1}
        
        for label, attr, val in [("选S:", "s_sel", "0.1"), ("粘P:", "s_pas", "0.2"), ("回E:", "s_ent", "0.8")]:
            tk.Label(params_f, text=label, font=("微软雅黑", 8, "bold"), bg="#90ee90").pack(side=tk.LEFT, padx=1)
            s = tk.Spinbox(params_f, **spin_opt)
            s.delete(0, "end"); s.insert(0, val)
            s.pack(side=tk.LEFT, padx=2)
            setattr(self, attr, s)

        # 按钮区
        tk.Button(params_f, text="🔥批量录入", command=self.pop_preview, bg="#ffffff", font=("微软雅黑", 8, "bold"), bd=1).pack(side=tk.RIGHT, padx=2)

        # 3. 集成滚动日志区
        self.log_text = tk.Text(self.main_f, font=("Consolas", 9), bg="#ffffff", fg="#2c3e50", bd=0, height=12)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=8, pady=2)
        self.log_text.tag_config("dup", background="#ffb2b2", foreground="#b22222")
        self.log_text.tag_config("auto", foreground="#2980b9") 

        # 4. 底部状态
        self.info_f = tk.Frame(self.main_f, bg="#90ee90")
        self.info_f.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Button(self.info_f, text="CLR", font=("Consolas", 7), command=self.clear_all, bd=0).pack(side=tk.LEFT, padx=8)
        self.info_lbl = tk.Label(self.info_f, text=f"Total: {len(BARCODE_HISTORY)}", font=("Consolas", 9, "bold"), bg="#90ee90")
        self.info_lbl.pack(side=tk.RIGHT, padx=8)
        
        self.title_bar.bind("<Button-1>", self.start_move); self.title_bar.bind("<B1-Motion>", self.do_move)
        self.log_count = len(BARCODE_HISTORY)

    def pop_preview(self):
        """弹出录入前核对窗"""
        try:
            raw = self.root.clipboard_get().split('\n')
            sns = sorted(list(set(s.strip() for s in raw if s.strip())))
            if not sns: return
        except: return

        self.pv = tk.Toplevel(self.root)
        self.pv.title("待录入清单 (Del键可删)")
        self.pv.geometry("260x380")
        self.pv.attributes("-topmost", True)
        
        self.lb = tk.Listbox(self.pv, font=("Consolas", 10), selectmode=tk.MULTIPLE)
        self.lb.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        for s in sns: self.lb.insert(tk.END, s)
        self.lb.bind("<Delete>", lambda e: [self.lb.delete(i) for i in reversed(self.lb.curselection())])

        tk.Button(self.pv, text="🚀 确认无误，开始录入", command=self.execute_auto, bg="#27ae60", fg="white", font=("微软雅黑", 9, "bold")).pack(fill=tk.X, padx=5, pady=5)

    def execute_auto(self):
        sns = list(self.lb.get(0, tk.END)); self.pv.destroy()
        if sns:
            self.root.attributes("-alpha", 0.4)
            threading.Thread(target=self._auto_run, args=(sns,), daemon=True).start()

    def trigger_alarm(self, is_dup):
        if is_dup: winsound.Beep(1200, 600)
        def flash(s):
            if s < 6 and is_dup:
                c = "#ffffff" if s % 2 == 0 else "#ff5252"
                self.main_f.config(bg=c); self.info_f.config(bg=c)
                self.root.after(250, lambda: flash(s + 1))
            else:
                final_bg = "#ffcccb" if is_dup else "#90ee90"
                self.main_f.config(bg=final_bg); self.info_f.config(bg=final_bg)
        flash(0)

    def add_log(self, code, status_tag=None):
        self.log_text.config(state=tk.NORMAL)
        ts = time.strftime("%H:%M:%S")
        self.log_count += 1
        tag = status_tag if status_tag else None
        status = {"dup": "DUP", "auto": "AUTO"}.get(status_tag, "OK")
        self.log_text.insert(tk.END, f"[{self.log_count:02d}] {ts} {status}: {code}\n", tag)
        self.log_text.see(tk.END); self.log_text.config(state=tk.DISABLED)
        self.info_lbl.config(text=f"Total: {len(BARCODE_HISTORY)}")

    def update_monitor(self, code, is_dup):
        self.trigger_alarm(is_dup)
        if is_dup:
            self.add_log(code, "dup")
            # 重复时拉回并全选
            with kb_controller.pressed(Key.shift): kb_controller.press(Key.tab); kb_controller.release(Key.tab)
            time.sleep(float(self.s_sel.get()))
            with kb_controller.pressed(Key.ctrl): kb_controller.press('a'); kb_controller.release('a')
        else:
            self.add_log(code)
            with open(HISTORY_FILE, "a") as f: f.write(f"{code}\n")

    def _auto_run(self, sns):
        time.sleep(4)
        ds, dp, de = float(self.s_sel.get()), float(self.s_pas.get()), float(self.s_ent.get())
        for sn in sns:
            with kb_controller.pressed(Key.ctrl): kb_controller.press('a'); kb_controller.release('a')
            time.sleep(ds)
            self.root.after(0, lambda x=sn: [self.root.clipboard_clear(), self.root.clipboard_append(x)])
            time.sleep(dp)
            with kb_controller.pressed(Key.ctrl): kb_controller.press('v'); kb_controller.release('v')
            time.sleep(0.1)
            kb_controller.press(Key.enter); kb_controller.release(Key.enter)
            if sn not in BARCODE_HISTORY:
                BARCODE_HISTORY.add(sn); with open(HISTORY_FILE, "a") as f: f.write(f"{sn}\n")
            self.root.after(0, self.add_log, sn, "auto")
            time.sleep(de)
        self.root.after(0, lambda: [self.root.attributes("-alpha", 0.92), winsound.Beep(1000, 300)])

    def start_move(self, event): self.x, self.y = event.x, event.y
    def do_move(self, event): self.root.geometry(f"+{self.root.winfo_x()+event.x-self.x}+{self.root.winfo_y()+event.y-self.y}")
    def minimize(self):
        self.root.overrideredirect(False); self.root.iconify()
        self.root.bind("<FocusIn>", lambda e: [self.root.overrideredirect(True), self.root.unbind("<FocusIn>")] )
    def clear_all(self):
        if messagebox.askyesno("全清确认", "警告：将彻底清除所有记录和文件！"):
            BARCODE_HISTORY.clear(); self.log_count = 0
            if os.path.exists(HISTORY_FILE): os.remove(HISTORY_FILE)
            self.log_text.config(state=tk.NORMAL); self.log_text.delete('1.0', tk.END); self.log_text.config(state=tk.DISABLED)
            self.info_lbl.config(text="Total: 0")
    def safe_exit(self): self.root.quit(); os._exit(0)

def on_press(key):
    global LAST_KEY_TIME, SCAN_BUFFER
    now = time.time(); interval = now - LAST_KEY_TIME; LAST_KEY_TIME = now
    try:
        if key == Key.enter:
            code = "".join(SCAN_BUFFER).strip()
            if code:
                is_dup = code in BARCODE_HISTORY
                if not is_dup: BARCODE_HISTORY.add(code)
                app.root.after(0, app.update_monitor, code, is_dup)
            SCAN_BUFFER = []
        elif hasattr(key, 'char') and key.char:
            if interval > SCAN_SPEED_THRESHOLD: SCAN_BUFFER = []
            SCAN_BUFFER.append(key.char)
    except: pass

if __name__ == "__main__":
    root = tk.Tk()
    app = UltimateThinMonitor(root)
    threading.Thread(target=lambda: keyboard.Listener(on_press=on_press).start(), daemon=True).start()
    root.mainloop()
