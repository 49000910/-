import time, threading, winsound, os
import tkinter as tk
from tkinter import messagebox, ttk
from pynput import keyboard, mouse
from pynput.keyboard import Controller, Key
from pynput.mouse import Controller as MouseController

# --- 核心配置 ---
HISTORY_FILE = "barcode_history.txt"
BARCODE_HISTORY = set()
SCAN_BUFFER = []
LAST_KEY_TIME = 0
SCAN_SPEED_THRESHOLD = 0.05 
kb, ms = Controller(), MouseController()

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
        self.root.geometry("260x245")
        self.root.attributes("-topmost", True, "-alpha", 0.96)
        self.root.overrideredirect(True)
        
        self.is_running_batch = False
        self.stop_batch = False
        self.batch_added = []
        self.sub = None
        self.last_x, self.last_y = 0, 0 

        self.themes = {
            "def": {"bg": "#ECEFF1", "head": "#CFD8DC", "title": "#90A4AE", "txt_bg": "#FFFFFF", "title_fg": "#37474F", "pb": "#90A4AE"},
            "ok":  {"bg": "#A5D6A7", "head": "#A5D6A7", "title": "#66BB6A", "txt_bg": "#E8F5E9", "title_fg": "#1B5E20", "pb": "#4CAF50"},
            "dup": {"bg": "#EF9A9A", "head": "#EF9A9A", "title": "#E57373", "txt_bg": "#FFEBEE", "title_fg": "#FFFFFF", "pb": "#F44336"}
        }

        # --- UI 构建 ---
        self.title_bar = tk.Frame(self.root, height=25)
        self.title_bar.pack(fill=tk.X)
        self.title_lbl = tk.Label(self.title_bar, text=" 🛡️ 采集助手 V5.4", font=("微软雅黑", 9, "bold"))
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
        tk.Button(self.ctrl_f, text="批量任务库", command=self.open_sub_win, bg="#CFD8DC", font=("微软雅黑", 8), relief=tk.FLAT).pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
        tk.Button(self.ctrl_f, text="清", command=self.clear_history, bg="#FFCCBC", fg="#D84315", font=("微软雅黑", 8, "bold"), relief=tk.FLAT, width=4).pack(side=tk.RIGHT, padx=2)

        self.ps = ttk.Style(); self.ps.theme_use('default')
        self.ps.configure("TProgressbar", thickness=4, bd=0, troughcolor="#E0E0E0", background="#4CAF50")
        self.p_bar = ttk.Progressbar(self.root, orient=tk.HORIZONTAL, mode='determinate', style="TProgressbar")

        self.log_text = tk.Text(self.root, font=("Consolas", 8), height=8, bd=0, padx=5)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self.log_text.tag_config("curr_txt", font=("Consolas", 11, "bold")) 
        self.log_text.tag_config("dup_txt", foreground="#C62828")
        self.log_text.tag_config("bat_txt", foreground="#1B5E20")
        self.log_text.tag_config("sys_txt", foreground="#E65100", font=("Consolas", 9, "bold"))

        self.info_lbl = tk.Label(self.root, text=f"Total: {len(BARCODE_HISTORY)}", font=("Arial", 7))
        self.info_lbl.pack(side=tk.RIGHT, padx=5)

        self.set_theme_color("def")
        keyboard.Listener(on_press=self.on_press).start()

    # --- 基础交互 ---
    def start_move(self, e): self.x, self.y = e.x, e.y
    def do_move(self, e): self.root.geometry(f"+{self.root.winfo_x()+(e.x-self.x)}+{self.root.winfo_y()+(e.y-self.y)}")
    def set_theme_color(self, key):
        if self.is_running_batch: return 
        t = self.themes[key]
        for w in [self.root, self.ctrl_f, self.params_f, self.info_lbl]: w.configure(bg=t["bg"])
        self.title_bar.configure(bg=t["title"]); self.title_lbl.configure(bg=t["title"], fg=t["title_fg"])
        self.log_text.configure(bg=t["txt_bg"]); self.ps.configure("TProgressbar", background=t["pb"])

    def fade_out_sub(self, window, alpha=0.98):
        if alpha > 0:
            alpha -= 0.2
            window.attributes("-alpha", alpha)
            self.root.after(15, lambda: self.fade_out_sub(window, alpha))
        else:
            window.destroy()
            self.sub = None

    def open_sub_win(self):
        if self.sub and self.sub.winfo_exists(): return
        self.sub = tk.Toplevel(self.root)
        self.sub.overrideredirect(True); self.sub.geometry("240x350")
        self.sub.attributes("-topmost", True, "-alpha", 0.98); self.sub.configure(bg="#F5F7F9")
        sub_t = tk.Frame(self.sub, bg="#455A64", height=25); sub_t.pack(fill=tk.X)
        tk.Label(sub_t, text=" 批量库 (双击删除单项)", fg="white", bg="#455A64", font=("微软雅黑", 8, "bold")).pack(side=tk.LEFT)
        tk.Button(sub_t, text="✕", command=lambda: self.fade_out_sub(self.sub), bg="#455A64", fg="#CFD8DC", bd=0, padx=8).pack(side=tk.RIGHT)
        sub_t.bind("<Button-1>", lambda e: setattr(self, 'sx', e.x) or setattr(self, 'sy', e.y))
        sub_t.bind("<B1-Motion>", lambda e: self.sub.geometry(f"+{self.sub.winfo_x()+(e.x-self.sx)}+{self.sub.winfo_y()+(e.y-self.sy)}"))
        btn_f = tk.Frame(self.sub, bg="#F5F7F9", pady=5); btn_f.pack(fill=tk.X, padx=5)
        tk.Button(btn_f, text="📋 粘贴排序", command=self.clip_load, bg="#E1F5FE", font=("微软雅黑", 8), bd=0).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        tk.Button(btn_f, text="🚀 开始执行", command=self.start_batch, bg="#E8F5E9", font=("微软雅黑", 8, "bold"), bd=0).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        lf = tk.Frame(self.sub, bg="white", bd=1, relief=tk.SOLID); lf.pack(fill=tk.BOTH, expand=True, padx=8, pady=5)
        self.listb = tk.Listbox(lf, font=("Consolas", 10), bd=0, highlightthickness=0); self.listb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = tk.Scrollbar(lf, width=8, command=self.listb.yview); sb.pack(side=tk.RIGHT, fill=tk.Y); self.listb.config(yscrollcommand=sb.set)
        self.listb.bind("<Double-Button-1>", lambda e: self.listb.delete(self.listb.curselection()))

    # --- 执行逻辑优化 ---
    def start_batch(self):
        codes = self.listb.get(0, tk.END)
        if not codes or self.is_running_batch: return
        
        # 1. 点击执行后，子窗口执行消失动画
        if self.sub: self.fade_out_sub(self.sub)
        
        self.is_running_batch = True; self.stop_batch = False; self.batch_added = []
        self.root.attributes("-alpha", 0.45) # 主窗变淡
        self.p_bar.pack(fill=tk.X, padx=2, before=self.log_text); self.p_bar['maximum'] = len(codes); self.p_bar['value'] = 0
        
        threading.Thread(target=self.prepare_and_run, args=(codes,), daemon=True).start()

    def prepare_and_run(self, codes):
        p = ms.position
        self.last_x, self.last_y = p[0], p[1]
        for i in range(5, 0, -1):
            self.root.after(0, lambda x=i: self.log_text.insert("1.0", f"⏳ 准备中... {x}s\n", "sys_txt"))
            for _ in range(10):
                if self.fast_panic_check(): 
                    self.root.after(0, self.abort_mission); return
                time.sleep(0.1)
        winsound.Beep(1500, 150)
        self.batch_engine(codes)

    def fast_panic_check(self):
        p = ms.position
        dist = abs(p[0] - self.last_x) + abs(p[1] - self.last_y)
        self.last_x, self.last_y = p[0], p[1]
        if dist > 35:
            self.stop_batch = True; return True
        return False

    def batch_engine(self, codes):
        e_delay, m_delay = float(self.spin_e1.get()), float(self.spin_mid.get())
        for idx, code in enumerate(codes):
            if self.fast_panic_check(): break
            kb.type(code); time.sleep(e_delay)
            if self.fast_panic_check(): break
            kb.tap(Key.enter)
            if self.r2_var.get(): time.sleep(0.05); kb.tap(Key.enter)
            self.root.after(0, self.update_ui, idx + 1, code)
            for _ in range(int(m_delay / 0.05)):
                if self.fast_panic_check(): break
                time.sleep(0.05)
            if self.stop_batch: break
        self.is_running_batch = False
        self.root.after(0, self.finalize_batch)

    def update_ui(self, val, code):
        self.p_bar['value'] = val; self.handle_scan(code, True)

    def abort_mission(self):
        self.log_text.insert("1.0", "🚫 任务已中止\n", "dup_txt")
        self.is_running_batch = False; self.p_bar.pack_forget(); self.root.attributes("-alpha", 0.96)

    def finalize_batch(self):
        self.root.attributes("-alpha", 0.96); self.p_bar.pack_forget()
        if self.stop_batch:
            if messagebox.askyesno("止付成功", "已停止录入。是否回滚？"):
                for c in self.batch_added:
                    if c in BARCODE_HISTORY: BARCODE_HISTORY.remove(c)
                with open(HISTORY_FILE, "w", encoding="utf-8") as f:
                    for c in BARCODE_HISTORY: f.write(c + "\n")
                self.info_lbl.config(text=f"Total: {len(BARCODE_HISTORY)}")
        else:
            winsound.Beep(1200, 200); messagebox.showinfo("完成", "录入任务结束")

    def clip_load(self):
        try:
            raw = self.root.clipboard_get()
            items = sorted(list(set([l.strip() for l in raw.split('\n') if l.strip()])))
            self.listb.delete(0, tk.END)
            for i in items: self.listb.insert(tk.END, i)
        except: pass

    def handle_scan(self, barcode, is_batch=False):
        self.log_text.tag_remove("curr_txt", "1.0", tk.END)
        if is_batch:
            if barcode not in BARCODE_HISTORY:
                BARCODE_HISTORY.add(barcode); self.batch_added.append(barcode)
                with open(HISTORY_FILE, "a", encoding="utf-8") as f: f.write(barcode + "\n")
            self.log_text.insert("1.0", f"● {barcode}\n", ("curr_txt", "bat_txt"))
        else:
            if barcode in BARCODE_HISTORY:
                winsound.Beep(1000, 400); self.set_theme_color("dup")
                self.log_text.insert("1.0", f"❌ {barcode}\n", ("curr_txt", "dup_txt"))
                if self.pb_var.get():
                    with kb.pressed(Key.shift): kb.tap(Key.tab)
                    time.sleep(0.02); 
                    with kb.pressed(Key.ctrl): kb.tap('a')
            else:
                self.set_theme_color("ok"); BARCODE_HISTORY.add(barcode)
                with open(HISTORY_FILE, "a", encoding="utf-8") as f: f.write(barcode + "\n")
                self.log_text.insert("1.0", f"✔ {barcode}\n", "curr_txt")
        self.info_lbl.config(text=f"Total: {len(BARCODE_HISTORY)}"); self.log_text.see("1.0")

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

    def clear_history(self):
        if messagebox.askyesno("确认", "清空？"):
            BARCODE_HISTORY.clear(); self.log_text.delete("1.0", tk.END); self.info_lbl.config(text="Total: 0")
            if os.path.exists(HISTORY_FILE): os.remove(HISTORY_FILE)

if __name__ == "__main__":
    root = tk.Tk(); app = UltimateMiniGuard(root); root.mainloop()
