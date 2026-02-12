import time, threading, winsound, os
import tkinter as tk
from tkinter import messagebox
from pynput import keyboard, mouse
from pynput.keyboard import Controller, Key
from pynput.mouse import Controller as MouseController

# --- 核心配置 ---
HISTORY_FILE = "barcode_history.txt"
BARCODE_HISTORY = set()
SCAN_BUFFER = []
LAST_KEY_TIME = 0
SCAN_SPEED_THRESHOLD = 0.05 
kb_controller = Controller()
ms_controller = MouseController()

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
        self.root.geometry("260x205") 
        self.root.attributes("-topmost", True)
        self.normal_alpha, self.work_alpha = 0.96, 0.45
        self.root.attributes("-alpha", self.normal_alpha)
        self.root.overrideredirect(True) 

        self.current_sns = [] 
        self.session_sns = [] 
        
        self.themes = {
            "def": {"bg": "#ECEFF1", "head": "#CFD8DC", "title": "#90A4AE", "txt_bg": "#FFFFFF", "title_fg": "#37474F"},
            "ok":  {"bg": "#A5D6A7", "head": "#A5D6A7", "title": "#66BB6A", "txt_bg": "#E8F5E9", "title_fg": "#1B5E20"},
            "dup": {"bg": "#EF9A9A", "head": "#EF9A9A", "title": "#E57373", "txt_bg": "#FFEBEE", "title_fg": "#FFFFFF"}
        }
        
        self.title_bar = tk.Frame(self.root, height=22)
        self.title_bar.pack(fill=tk.X)
        self.title_lbl = tk.Label(self.title_bar, text=" 采集助手", font=("微软雅黑", 8, "bold"))
        self.title_lbl.pack(side=tk.LEFT)
        tk.Button(self.title_bar, text="✕", command=root.quit, bg="#FF7043", fg="white", font=("Arial", 7, "bold"), bd=0, padx=5).pack(side=tk.RIGHT)

        for widget in [self.title_bar, self.title_lbl, self.root]:
            widget.bind("<Button-1>", self.start_move)
            widget.bind("<B1-Motion>", self.do_move)

        self.params_f = tk.Frame(self.root, pady=1)
        self.params_f.pack(fill=tk.X)
        spin_opt = {"font": ("Consolas", 8), "width": 3, "from_": 0.0, "to": 5.0, "increment": 0.05}
        
        self.pb_var = tk.BooleanVar(value=True)
        self.cb_pb = tk.Checkbutton(self.params_f, text="PB", variable=self.pb_var, font=("微软雅黑", 8))
        self.cb_pb.pack(side=tk.LEFT)
        self.r2_var = tk.BooleanVar(value=False)
        self.cb_r2 = tk.Checkbutton(self.params_f, text="回2", variable=self.r2_var, font=("微软雅黑", 8))
        self.cb_r2.pack(side=tk.LEFT)
        
        tk.Label(self.params_f, text="E:", font=("微软雅黑", 8)).pack(side=tk.LEFT)
        self.spin_e1 = tk.Spinbox(self.params_f, **spin_opt)
        self.spin_e1.delete(0, "end"); self.spin_e1.insert(0, "0.01"); self.spin_e1.pack(side=tk.LEFT)

        tk.Label(self.params_f, text="待:", font=("微软雅黑", 8)).pack(side=tk.LEFT)
        self.spin_mid = tk.Spinbox(self.params_f, **spin_opt)
        self.spin_mid.delete(0, "end"); self.spin_mid.insert(0, "0.85"); self.spin_mid.pack(side=tk.LEFT)
        
        self.ctrl_f = tk.Frame(self.root, pady=1)
        self.ctrl_f.pack(fill=tk.X)
        tk.Button(self.ctrl_f, text="批量录入窗", command=self.open_sub_win, bg="#CFD8DC", font=("微软雅黑", 8), relief=tk.FLAT).pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
        tk.Button(self.ctrl_f, text="清", command=self.clear_history, bg="#FFCCBC", fg="#D84315", font=("微软雅黑", 8, "bold"), relief=tk.FLAT, width=3).pack(side=tk.RIGHT, padx=2)

        self.log_text = tk.Text(self.root, font=("Consolas", 8), height=7, bd=0, padx=5)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=1)
        self.log_text.tag_config("curr_txt", font=("Consolas", 10, "bold")) 
        self.log_text.tag_config("dup_txt", foreground="#C62828")
        self.log_text.tag_config("bat_txt", foreground="#1B5E20")
        self.log_text.tag_config("sys_txt", foreground="#E65100", font=("Consolas", 8, "bold"))

        self.info_lbl = tk.Label(self.root, text=f"Cnt: {len(BARCODE_HISTORY)}", font=("Arial", 7))
        self.info_lbl.pack(side=tk.RIGHT, padx=2)

        self.set_theme_color("def")
        self.listener = keyboard.Listener(on_press=self.on_press); self.listener.start()

    def start_move(self, event): self.x = event.x; self.y = event.y
    def do_move(self, event):
        self.root.geometry(f"+{self.root.winfo_x()+(event.x-self.x)}+{self.root.winfo_y()+(event.y-self.y)}")

    def set_theme_color(self, type_key):
        t = self.themes[type_key]
        for w in [self.root, self.ctrl_f, self.info_lbl]: w.configure(bg=t["bg"])
        self.title_bar.configure(bg=t["title"])
        self.title_lbl.configure(bg=t["title"], fg=t["title_fg"])
        self.params_f.configure(bg=t["head"])
        for w in [self.cb_pb, self.cb_r2]: w.configure(bg=t["head"], activebackground=t["head"])
        self.log_text.configure(bg=t["txt_bg"])

    def clear_history(self):
        if messagebox.askyesno("确认", "清空所有采集历史？"):
            BARCODE_HISTORY.clear()
            if os.path.exists(HISTORY_FILE): os.remove(HISTORY_FILE)
            self.set_theme_color("def"); self.log_text.delete("1.0", tk.END)
            self.info_lbl.config(text="Cnt: 0")

    def handle_scan(self, barcode, is_batch=False):
        self.log_text.tag_remove("curr_txt", "1.0", tk.END)
        if is_batch:
            if barcode not in BARCODE_HISTORY:
                BARCODE_HISTORY.add(barcode)
                with open(HISTORY_FILE, "a", encoding="utf-8") as f: f.write(barcode + "\n")
            self.log_text.insert("1.0", f"[批] {barcode}\n", ("curr_txt", "bat_txt"))
            self.info_lbl.config(text=f"Cnt: {len(BARCODE_HISTORY)}")
        else:
            if barcode in BARCODE_HISTORY:
                winsound.Beep(1000, 300); self.set_theme_color("dup")
                self.log_text.insert("1.0", f"[重] {barcode}\n", ("curr_txt", "dup_txt"))
                if self.pb_var.get():
                    with kb_controller.pressed(Key.shift):
                        kb_controller.press(Key.tab)
                        kb_controller.release(Key.tab)
                    time.sleep(0.01)
                    with kb_controller.pressed(Key.ctrl):
                        kb_controller.press('a')
                        kb_controller.release('a')
            else:
                self.set_theme_color("ok")
                BARCODE_HISTORY.add(barcode)
                with open(HISTORY_FILE, "a", encoding="utf-8") as f: f.write(barcode + "\n")
                self.log_text.insert("1.0", f"[OK] {barcode}\n", "curr_txt")
                self.info_lbl.config(text=f"Cnt: {len(BARCODE_HISTORY)}")
        self.log_text.see("1.0")

    def on_press(self, key):
        global LAST_KEY_TIME, SCAN_BUFFER
        now = time.time(); interval = now - LAST_KEY_TIME; LAST_KEY_TIME = now
        try:
            char = key.char if hasattr(key, 'char') and key.char else ('\n' if key == Key.enter else None)
            if not char: return
            if interval < SCAN_SPEED_THRESHOLD:
                if char == '\n':
                    bc = "".join(SCAN_BUFFER).strip(); SCAN_BUFFER = []
                    if bc: self.root.after(0, self.handle_scan, bc)
                else: SCAN_BUFFER.append(char)
            else: SCAN_BUFFER = [char] if char != '\n' else []
        except: pass

    # --- 批量录入相关逻辑 ---
    def open_sub_win(self):
        self.sub_win = tk.Toplevel(self.root)
        self.sub_win.title("批量录入")
        self.sub_win.geometry("240x320")
        self.sub_win.attributes("-topmost", True)
        
        f = tk.Frame(self.sub_win, bg="#f0f0f0")
        f.pack(fill=tk.X, pady=1)
        tk.Button(f, text="1. 读取剪贴板", command=self.load_from_clipboard, bg="#B3E5FC", font=("微软雅黑", 8)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=1)
        tk.Button(f, text="2. 执行录入", command=self.start_batch_input, bg="#C8E6C9", font=("微软雅黑", 8, "bold")).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=1)
        tk.Button(f, text="清", command=self.clear_sub_list, bg="#FFCCBC", font=("微软雅黑", 8), width=4).pack(side=tk.RIGHT, padx=1)
        
        self.listbox = tk.Listbox(self.sub_win, font=("Consolas", 9), bd=1)
        self.listbox.pack(fill=tk.BOTH, expand=True)

    def load_from_clipboard(self):
        try:
            data = self.root.clipboard_get()
            lines = [line.strip() for line in data.split('\n') if line.strip()]
            self.listbox.delete(0, tk.END)
            for line in lines: self.listbox.insert(tk.END, line)
        except: messagebox.showwarning("提示", "剪贴板无文本内容")

    def clear_sub_list(self): self.listbox.delete(0, tk.END)

    def start_batch_input(self):
        items = self.listbox.get(0, tk.END)
        if not items: return
        self.root.attributes("-alpha", self.work_alpha) # 变淡表示工作中
        threading.Thread(target=self.batch_worker, args=(items,), daemon=True).start()

    def batch_worker(self, items):
        e1 = float(self.spin_e1.get())
        mid = float(self.spin_mid.get())
        for code in items:
            # 模拟手动输入
            kb_controller.type(code)
            time.sleep(e1)
            kb_controller.press(Key.enter); kb_controller.release(Key.enter)
            
            # 更新主界面UI
            self.root.after(0, self.handle_scan, code, True)
            
            # 模拟"回2"逻辑
            if self.r2_var.get():
                time.sleep(0.05)
                kb_controller.press(Key.enter); kb_controller.release(Key.enter)
            
            time.sleep(mid) # 待机等待
        
        self.root.after(0, lambda: self.root.attributes("-alpha", self.normal_alpha))
        messagebox.showinfo("完成", f"批量录入结束，共计 {len(items)} 条")

if __name__ == "__main__":
    root = tk.Tk()
    app = UltimateMiniGuard(root)
    root.mainloop()
