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
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            BARCODE_HISTORY = set(line.strip() for line in f if line.strip())
    except: pass

class UltraThinMonitor:
    def __init__(self, root):
        self.root = root
        self.root.title("智能助手 v4.8")
        self.root.geometry("400x230") # 宽度微调以容纳所有横排开关
        self.root.attributes("-topmost", True, "-alpha", 0.92) 
        self.root.overrideredirect(True)

        # 1. 标题栏
        self.title_bar = tk.Frame(self.root, bg="#2c3e50", height=20)
        self.title_bar.pack(fill=tk.X)
        tk.Label(self.title_bar, text=" 🛡️ Guard v4.8 (级简全功能)", font=("微软雅黑", 8, "bold"), fg="white", bg="#2c3e50").pack(side=tk.LEFT)
        tk.Button(self.title_bar, text="×", bg="#2c3e50", fg="white", bd=0, command=self.safe_exit, font=("Arial", 8)).pack(side=tk.RIGHT, padx=5)
        tk.Button(self.title_bar, text="—", bg="#2c3e50", fg="white", bd=0, command=self.minimize, font=("Arial", 8)).pack(side=tk.RIGHT, padx=5)

        # 2. 交互主区
        self.main_f = tk.Frame(self.root, bg="#90ee90", pady=2) 
        self.main_f.pack(fill=tk.BOTH, expand=True)

        # --- 核心横排参数行 (PB开关 | E1 | E2开关 | E2 | 批量) ---
        params_f = tk.Frame(self.main_f, bg="#90ee90")
        params_f.pack(fill=tk.X, padx=2, pady=2)
        spin_opt = {"font": ("Consolas", 9), "width": 3, "from_": 0.0, "to": 5.0, "increment": 0.1, "bd": 1}
        
        # [开关] 拦截重复拉回 PB (PullBack)
        self.use_pb = tk.BooleanVar(value=True)
        tk.Checkbutton(params_f, variable=self.use_pb, bg="#90ee90", activebackground="#90ee90", bd=0).pack(side=tk.LEFT)
        tk.Label(params_f, text="PB", font=("Consolas", 8, "bold"), bg="#90ee90").pack(side=tk.LEFT)

        # [数字] E1
        tk.Label(params_f, text="E1:", font=("Consolas", 8, "bold"), bg="#90ee90").pack(side=tk.LEFT, padx=(2,0))
        self.spin_e1 = tk.Spinbox(params_f, **spin_opt)
        self.spin_e1.delete(0, "end"); self.spin_e1.insert(0, "0.2")
        self.spin_e1.pack(side=tk.LEFT, padx=1)

        # [开关] 双回车
        self.use_double_enter = tk.BooleanVar(value=True)
        tk.Checkbutton(params_f, variable=self.use_double_enter, bg="#90ee90", activebackground="#90ee90", bd=0).pack(side=tk.LEFT, padx=(2,0))
        
        # [数字] E2
        tk.Label(params_f, text="E2:", font=("Consolas", 8, "bold"), bg="#90ee90").pack(side=tk.LEFT)
        self.spin_e2 = tk.Spinbox(params_f, **spin_opt)
        self.spin_e2.delete(0, "end"); self.spin_e2.insert(0, "0.8")
        self.spin_e2.pack(side=tk.LEFT, padx=1)

        # [按钮] 批量
        tk.Button(params_f, text="🔥批量", command=self.pop_preview, bg="#ffffff", font=("微软雅黑", 8, "bold"), bd=1, padx=2).pack(side=tk.RIGHT, padx=2)

        # 3. 日志区
        self.log_text = tk.Text(self.main_f, font=("Consolas", 9), bg="#ffffff", fg="#2c3e50", bd=0, height=8)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=2)
        self.log_text.tag_config("dup", background="#ffb2b2", foreground="#b22222")
        self.log_text.tag_config("auto", foreground="#2980b9") 

        # 4. 底部状态
        self.info_f = tk.Frame(self.main_f, bg="#90ee90")
        self.info_f.pack(fill=tk.X, side=tk.BOTTOM)
        self.info_lbl = tk.Label(self.info_f, text=f"Total:{len(BARCODE_HISTORY)}", font=("Consolas", 8, "bold"), bg="#90ee90")
        self.info_lbl.pack(side=tk.RIGHT, padx=5)
        
        self.title_bar.bind("<Button-1>", self.start_move); self.title_bar.bind("<B1-Motion>", self.do_move)
        self.log_count = len(BARCODE_HISTORY)

    def pop_preview(self):
        try:
            raw = self.root.clipboard_get().split('\n')
            sns = sorted(list(set(s.strip() for s in raw if s.strip())))
            if not sns: return
        except: return
        pv = tk.Toplevel(self.root)
        pv.title("核对"); pv.geometry("240x350"); pv.attributes("-topmost", True)
        lb = tk.Listbox(pv, font=("Consolas", 10), selectmode=tk.MULTIPLE)
        lb.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        for s in sns: lb.insert(tk.END, s)
        lb.bind("<Delete>", lambda e: [lb.delete(i) for i in reversed(lb.curselection())])
        btn_f = tk.Frame(pv); btn_f.pack(fill=tk.X, pady=5)
        tk.Button(btn_f, text="🗑️删除", command=lambda: [lb.delete(i) for i in reversed(lb.curselection())], fg="#c0392b").pack(side=tk.LEFT, padx=10, expand=True, fill=tk.X)
        tk.Button(btn_f, text="🚀执行", command=lambda: self.execute_auto(pv, lb), bg="#27ae60", fg="white", font=("微软雅黑", 9, "bold")).pack(side=tk.LEFT, padx=10, expand=True, fill=tk.X)

    def execute_auto(self, pv, lb):
        sns = list(lb.get(0, tk.END)); pv.destroy()
        if sns:
            self.root.attributes("-alpha", 0.4)
            threading.Thread(target=self._auto_run, args=(sns,), daemon=True).start()

    def _auto_run(self, sns):
        time.sleep(4)
        e1, e2 = float(self.spin_e1.get()), float(self.spin_e2.get())
        double_mode = self.use_double_enter.get()
        for sn in sns:
            with kb_controller.pressed(Key.ctrl): kb_controller.press('a'); kb_controller.release('a')
            time.sleep(0.1) 
            self.root.after(0, lambda x=sn: [self.root.clipboard_clear(), self.root.clipboard_append(x)])
            time.sleep(e1)
            with kb_controller.pressed(Key.ctrl): kb_controller.press('v'); kb_controller.release('v')
            time.sleep(0.1)
            kb_controller.press(Key.enter); kb_controller.release(Key.enter)
            if double_mode:
                time.sleep(0.1); kb_controller.press(Key.enter); kb_controller.release(Key.enter); time.sleep(e2)
            else: time.sleep(0.4)
            if sn not in BARCODE_HISTORY:
                BARCODE_HISTORY.add(sn)
                with open(HISTORY_FILE, "a", encoding="utf-8") as f: f.write(f"{sn}\n")
            self.root.after(0, self.add_log, sn, "auto")
        self.root.after(0, lambda: [self.root.attributes("-alpha", 0.92), winsound.Beep(1000, 300)])

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
        self.info_lbl.config(text=f"Total:{len(BARCODE_HISTORY)}")

    def update_monitor(self, code, is_dup):
        self.trigger_alarm(is_dup)
        if is_dup:
            self.add_log(code, "dup")
            # 只有勾选 PB 开关时才拉回
            if self.use_pb.get():
                with kb_controller.pressed(Key.shift): kb_controller.press(Key.tab); kb_controller.release(Key.tab)
                time.sleep(0.2); with kb_controller.pressed(Key.ctrl): kb_controller.press('a'); kb_controller.release('a')
        else:
            self.add_log(code)
            with open(HISTORY_FILE, "a", encoding="utf-8") as f: f.write(f"{code}\n")

    def start_move(self, event): self.x, self.y = event.x, event.y
    def do_move(self, event): self.root.geometry(f"+{self.root.winfo_x()+event.x-self.x}+{self.root.winfo_y()+event.y-self.y}")
    def minimize(self):
        self.root.overrideredirect(False); self.root.iconify()
        self.root.bind("<FocusIn>", lambda e: [self.root.overrideredirect(True), self.root.unbind("<FocusIn>")] )
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
    app = UltraThinMonitor(root)
    threading.Thread(target=lambda: keyboard.Listener(on_press=on_press).start(), daemon=True).start()
    root.mainloop()
