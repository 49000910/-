import time, threading, winsound, os
import tkinter as tk
from tkinter import messagebox
from pynput import keyboard, mouse
from pynput.keyboard import Controller, Key
from pynput.mouse import Controller as MouseController

# --- 全局配置 ---
HISTORY_FILE = "barcode_history.txt"
BARCODE_HISTORY = set()
SCAN_BUFFER = []
LAST_KEY_TIME = 0
SCAN_SPEED_THRESHOLD = 0.05 
kb = Controller()
ms = MouseController()

def load_history():
    BARCODE_HISTORY.clear()
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    if line.strip(): BARCODE_HISTORY.add(line.strip())
        except: pass

load_history()

class UltimateMiniGuard:
    def __init__(self, root):
        self.root = root
        self.root.title("采集助手")
        self.root.geometry("260x220")
        self.root.attributes("-topmost", True, "-alpha", 0.96)
        self.root.overrideredirect(True)
        
        self.is_running_batch = False # 批量运行状态
        self.stop_batch = False      # 强行停止标志
        self.batch_added = []        # 本次批量新增的条码（用于回滚）

        # 样式定义
        self.themes = {
            "def": {"bg": "#ECEFF1", "head": "#CFD8DC", "title": "#90A4AE", "txt_bg": "#FFFFFF", "title_fg": "#37474F"},
            "ok":  {"bg": "#A5D6A7", "head": "#A5D6A7", "title": "#66BB6A", "txt_bg": "#E8F5E9", "title_fg": "#1B5E20"},
            "dup": {"bg": "#EF9A9A", "head": "#EF9A9A", "title": "#E57373", "txt_bg": "#FFEBEE", "title_fg": "#FFFFFF"}
        }

        # UI 构建
        self.title_bar = tk.Frame(self.root, height=25)
        self.title_bar.pack(fill=tk.X)
        self.title_lbl = tk.Label(self.title_bar, text=" 🛡️ 采集助手 V2", font=("微软雅黑", 9, "bold"))
        self.title_lbl.pack(side=tk.LEFT)
        tk.Button(self.title_bar, text="✕", command=root.quit, bg="#FF7043", fg="white", font=("Arial", 8, "bold"), bd=0, padx=8).pack(side=tk.RIGHT)

        for w in [self.title_bar, self.title_lbl]:
            w.bind("<Button-1>", self.start_move); w.bind("<B1-Motion>", self.do_move)

        self.params_f = tk.Frame(self.root, pady=2)
        self.params_f.pack(fill=tk.X)
        spin_opt = {"font": ("Consolas", 8), "width": 4, "from_": 0.0, "to": 9.9, "increment": 0.05}
        
        self.pb_var = tk.BooleanVar(value=True)
        tk.Checkbutton(self.params_f, text="PB", variable=self.pb_var, font=("微软雅黑", 8)).pack(side=tk.LEFT)
        self.r2_var = tk.BooleanVar(value=False)
        tk.Checkbutton(self.params_f, text="回2", variable=self.r2_var, font=("微软雅黑", 8)).pack(side=tk.LEFT)
        
        tk.Label(self.params_f, text="E:", font=("微软雅黑", 8)).pack(side=tk.LEFT)
        self.spin_e1 = tk.Spinbox(self.params_f, **spin_opt)
        self.spin_e1.delete(0, "end"); self.spin_e1.insert(0, "0.01"); self.spin_e1.pack(side=tk.LEFT)

        tk.Label(self.params_f, text="待:", font=("微软雅黑", 8)).pack(side=tk.LEFT)
        self.spin_mid = tk.Spinbox(self.params_f, **spin_opt)
        self.spin_mid.delete(0, "end"); self.spin_mid.insert(0, "0.85"); self.spin_mid.pack(side=tk.LEFT)
        
        self.ctrl_f = tk.Frame(self.root, pady=1)
        self.ctrl_f.pack(fill=tk.X)
        tk.Button(self.ctrl_f, text="批量录入窗", command=self.open_sub_win, bg="#CFD8DC", font=("微软雅黑", 8), relief=tk.FLAT).pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
        tk.Button(self.ctrl_f, text="清", command=self.clear_history, bg="#FFCCBC", fg="#D84315", font=("微软雅黑", 8, "bold"), relief=tk.FLAT, width=4).pack(side=tk.RIGHT, padx=2)

        self.log_text = tk.Text(self.root, font=("Consolas", 8), height=8, bd=0, padx=5)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self.log_text.tag_config("curr_txt", font=("Consolas", 11, "bold")) # 大一号字体
        self.log_text.tag_config("dup_txt", foreground="#C62828")
        self.log_text.tag_config("bat_txt", foreground="#1B5E20")

        self.info_lbl = tk.Label(self.root, text=f"Total: {len(BARCODE_HISTORY)}", font=("Arial", 7))
        self.info_lbl.pack(side=tk.RIGHT, padx=5)

        self.set_theme_color("def")
        keyboard.Listener(on_press=self.on_press).start()

    # --- 基础交互 ---
    def start_move(self, e): self.x, self.y = e.x, e.y
    def do_move(self, e): self.root.geometry(f"+{self.root.winfo_x()+(e.x-self.x)}+{self.root.winfo_y()+(e.y-self.y)}")

    def set_theme_color(self, key):
        if self.is_running_batch: return # 批量模式不闪烁变色
        t = self.themes[key]
        for w in [self.root, self.ctrl_f, self.info_lbl, self.params_f]: w.configure(bg=t["bg"])
        self.title_bar.configure(bg=t["title"]); self.title_lbl.configure(bg=t["title"], fg=t["title_fg"])
        self.log_text.configure(bg=t["txt_bg"])

    def clear_history(self):
        if messagebox.askyesno("确认", "清空所有本地记录？"):
            BARCODE_HISTORY.clear()
            if os.path.exists(HISTORY_FILE): os.remove(HISTORY_FILE)
            self.log_text.delete("1.0", tk.END); self.set_theme_color("def")
            self.info_lbl.config(text="Total: 0")

    # --- 核心扫码逻辑 ---
    def handle_scan(self, barcode, is_batch=False):
        self.log_text.tag_remove("curr_txt", "1.0", tk.END)
        if is_batch:
            # 批量逻辑：静默存储，不闪烁
            if barcode not in BARCODE_HISTORY:
                BARCODE_HISTORY.add(barcode); self.batch_added.append(barcode)
                with open(HISTORY_FILE, "a") as f: f.write(barcode + "\n")
            self.log_text.insert("1.0", f"● {barcode}\n", ("curr_txt", "bat_txt"))
        else:
            # 手动逻辑：重复拦截
            if barcode in BARCODE_HISTORY:
                winsound.Beep(1000, 400); self.set_theme_color("dup")
                self.log_text.insert("1.0", f"❌ {barcode}\n", ("curr_txt", "dup_txt"))
                if self.pb_var.get():
                    with kb.pressed(Key.shift): kb.tap(Key.tab)
                    time.sleep(0.02)
                    with kb.pressed(Key.ctrl): kb.tap('a')
            else:
                self.set_theme_color("ok")
                BARCODE_HISTORY.add(barcode)
                with open(HISTORY_FILE, "a") as f: f.write(barcode + "\n")
                self.log_text.insert("1.0", f"✔ {barcode}\n", "curr_txt")
        self.info_lbl.config(text=f"Total: {len(BARCODE_HISTORY)}")
        self.log_text.see("1.0")

    def on_press(self, key):
        global LAST_KEY_TIME, SCAN_BUFFER
        now = time.time(); interval = now - LAST_KEY_TIME; LAST_KEY_TIME = now
        try:
            c = key.char if hasattr(key, 'char') and key.char else ('\n' if key == Key.enter else None)
            if not c: return
            if interval < SCAN_SPEED_THRESHOLD:
                if c == '\n':
                    bc = "".join(SCAN_BUFFER).strip(); SCAN_BUFFER = []
                    if bc: self.root.after(0, self.handle_scan, bc)
                else: SCAN_BUFFER.append(c)
            else: SCAN_BUFFER = [c] if c != '\n' else []
        except: pass

    # --- 批量录入窗 ---
    def open_sub_win(self):
        self.sub = tk.Toplevel(self.root); self.sub.title("批量库"); self.sub.geometry("220x300"); self.sub.attributes("-topmost", True)
        btn_f = tk.Frame(self.sub)
        btn_f.pack(fill=tk.X)
        tk.Button(btn_f, text="📋 粘贴并排序", command=self.clip_load, bg="#B3E5FC", font=("微软雅黑", 8)).pack(side=tk.LEFT, expand=True, fill=tk.X)
        tk.Button(btn_f, text="🚀 开始", command=self.start_batch, bg="#C8E6C9", font=("微软雅黑", 8, "bold")).pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.listb = tk.Listbox(self.sub, font=("Consolas", 9)); self.listb.pack(fill=tk.BOTH, expand=True)

    def clip_load(self):
        try:
            raw = self.root.clipboard_get()
            # 自动去重并排序 (0-9, A-Z)
            items = sorted(list(set([l.strip() for l in raw.split('\n') if l.strip()])))
            self.listb.delete(0, tk.END)
            for i in items: self.listb.insert(tk.END, i)
        except: pass

    def start_batch(self):
        codes = self.listb.get(0, tk.END)
        if not codes or self.is_running_batch: return
        self.is_running_batch = True; self.stop_batch = False; self.batch_added = []
        self.root.attributes("-alpha", 0.45) # 避让透明度
        threading.Thread(target=self.batch_engine, args=(codes,), daemon=True).start()

    def batch_engine(self, codes):
        e_delay = float(self.spin_e1.get()); m_delay = float(self.spin_mid.get())
        last_pos = ms.position
        
        for code in codes:
            # 安全逻辑：鼠标快速移动则切断
            curr_pos = ms.position
            if abs(curr_pos[0]-last_pos[0]) > 50 or abs(curr_pos[1]-last_pos[1]) > 50:
                self.stop_batch = True; break
            last_pos = curr_pos

            # 模拟输入
            kb.type(code); time.sleep(e_delay)
            kb.tap(Key.enter)
            if self.r2_var.get(): time.sleep(0.05); kb.tap(Key.enter)
            
            self.root.after(0, self.handle_scan, code, True)
            time.sleep(m_delay)
        
        self.is_running_batch = False
        self.root.after(0, self.finalize_batch)

    def finalize_batch(self):
        self.root.attributes("-alpha", 0.96)
        if self.stop_batch:
            if messagebox.askyesno("录入中断", "检测到鼠标晃动！已停止。\n是否撤销本次已录入的记录？"):
                for c in self.batch_added:
                    if c in BARCODE_HISTORY: BARCODE_HISTORY.remove(c)
                # 重新写入文件（移除撤销的部分）
                with open(HISTORY_FILE, "w") as f:
                    for c in BARCODE_HISTORY: f.write(c + "\n")
                self.log_text.insert("1.0", "⚠️ 本次操作已回滚\n", "dup_txt")
                self.info_lbl.config(text=f"Total: {len(BARCODE_HISTORY)}")
        else:
            messagebox.showinfo("成功", "批量录入任务完成！")

if __name__ == "__main__":
    root = tk.Tk(); app = UltimateMiniGuard(root); root.mainloop()
