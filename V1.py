import time, threading, winsound, os
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

        self.current_sns = [] # 待录入列表

        self.clr_title_bar, self.clr_head = "#80CBC4", "#B2DFDB"
        self.clr_default_bg = "#ECEFF1"
        self.clr_ok, self.clr_ok_log = "#A5D6A7", "#E8F5E9"
        self.clr_dup, self.clr_dup_log = "#EF9A9A", "#FFEBEE"
        
        self.root.configure(bg=self.clr_default_bg)

        # --- 标题栏 ---
        self.title_bar = tk.Frame(self.root, bg=self.clr_title_bar, height=22)
        self.title_bar.pack(fill=tk.X)
        tk.Label(self.title_bar, text=" 采集助手", fg="#004D40", bg=self.clr_title_bar, font=("微软雅黑", 8, "bold")).pack(side=tk.LEFT)
        tk.Button(self.title_bar, text="✕", command=root.quit, bg="#FF7043", fg="white", font=("Arial", 7, "bold"), bd=0, padx=5).pack(side=tk.RIGHT)

        for widget in [self.title_bar, self.root]:
            widget.bind("<Button-1>", self.start_move)
            widget.bind("<B1-Motion>", self.do_move)

        # --- 参数区 ---
        self.params_f = tk.Frame(self.root, bg=self.clr_head, pady=1)
        self.params_f.pack(fill=tk.X)
        spin_opt = {"font": ("Consolas", 8), "width": 3, "from_": 0.0, "to": 5.0, "increment": 0.05}
        
        self.use_pb = tk.BooleanVar(value=True)
        tk.Checkbutton(self.params_f, text="PB", variable=self.use_pb, bg=self.clr_head, font=("微软雅黑", 8)).pack(side=tk.LEFT)
        self.use_double_enter = tk.BooleanVar(value=False)
        tk.Checkbutton(self.params_f, text="回2", variable=self.use_double_enter, bg=self.clr_head, font=("微软雅黑", 8)).pack(side=tk.LEFT)
        
        tk.Label(self.params_f, text="E:", bg=self.clr_head, font=("微软雅黑", 8)).pack(side=tk.LEFT)
        self.spin_e1 = tk.Spinbox(self.params_f, **spin_opt)
        self.spin_e1.delete(0, "end"); self.spin_e1.insert(0, "0.01"); self.spin_e1.pack(side=tk.LEFT)

        tk.Label(self.params_f, text="待:", bg=self.clr_head, font=("微软雅黑", 8)).pack(side=tk.LEFT)
        self.spin_mid = tk.Spinbox(self.params_f, **spin_opt)
        self.spin_mid.delete(0, "end"); self.spin_mid.insert(0, "0.85"); self.spin_mid.pack(side=tk.LEFT)
        
        # --- 控制区 ---
        self.ctrl_f = tk.Frame(self.root, bg=self.clr_default_bg)
        self.ctrl_f.pack(fill=tk.X, pady=1)
        tk.Button(self.ctrl_f, text="批量录入窗", command=self.open_sub_win, bg="#CFD8DC", font=("微软雅黑", 8), relief=tk.FLAT).pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
        tk.Button(self.ctrl_f, text="清", command=self.clear_history, bg="#FFCCBC", fg="#D84315", font=("微软雅黑", 8, "bold"), relief=tk.FLAT, width=3).pack(side=tk.RIGHT, padx=2)

        # --- 日志 ---
        self.log_text = tk.Text(self.root, font=("Consolas", 8), bg="#ffffff", height=7, bd=0, padx=5)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=1)
        self.log_text.tag_config("dup_text", foreground="#C62828", font=("Consolas", 8, "bold"))
        self.log_text.tag_config("batch_text", foreground="#1B5E20")

        self.info_lbl = tk.Label(self.root, text=f"Cnt: {len(BARCODE_HISTORY)}", font=("Arial", 7), bg=self.clr_default_bg)
        self.info_lbl.pack(side=tk.RIGHT, padx=2)

        self.listener = keyboard.Listener(on_press=self.on_press); self.listener.start()

    def start_move(self, event): self.x = event.x; self.y = event.y
    def do_move(self, event):
        self.root.geometry(f"+{self.root.winfo_x()+(event.x-self.x)}+{self.root.winfo_y()+(event.y-self.y)}")

    def clear_history(self):
        if messagebox.askyesno("确认", "清空所有已采集历史记录？"):
            BARCODE_HISTORY.clear()
            if os.path.exists(HISTORY_FILE): os.remove(HISTORY_FILE)
            self.log_text.configure(bg="#ffffff")
            self.log_text.insert("1.0", "[系统] 历史已清空\n")
            self.info_lbl.config(text="Cnt: 0")

    def handle_scan(self, barcode, is_batch=False):
        if barcode in BARCODE_HISTORY:
            if not is_batch: 
                winsound.Beep(1000, 300)
                self.root.configure(bg=self.clr_dup)
                self.log_text.configure(bg=self.clr_dup_log)
                self.info_lbl.configure(bg=self.clr_dup)
                self.log_text.insert("1.0", f"[重] {barcode}\n", "dup_text")
                if self.use_pb.get():
                    with kb_controller.pressed(Key.shift):
                        kb_controller.press(Key.tab)
                        kb_controller.release(Key.tab)
                    time.sleep(0.01)
                    with kb_controller.pressed(Key.ctrl):
                        kb_controller.press('a')
                        kb_controller.release('a')
        else:
            self.root.configure(bg=self.clr_ok)
            self.log_text.configure(bg=self.clr_ok_log)
            self.info_lbl.configure(bg=self.clr_ok)
            BARCODE_HISTORY.add(barcode)
            with open(HISTORY_FILE, "a", encoding="utf-8") as f: f.write(barcode + "\n")
            pfx, tag = ("[批]", "batch_text") if is_batch else ("[OK]", None)
            self.log_text.insert("1.0", f"{pfx} {barcode}\n", tag)
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

    def open_sub_win(self):
        self.sub_win = tk.Toplevel(self.root)
        self.sub_win.title("录入器")
        self.sub_win.geometry("240x320")
        self.sub_win.attributes("-topmost", True)
        btn_f = tk.Frame(self.sub_win, bg="#f0f0f0")
        btn_f.pack(fill=tk.X, pady=1)
        tk.Button(btn_f, text="1. 读取剪贴板", command=self.load_from_clipboard, bg="#B3E5FC", font=("微软雅黑", 8)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=1)
        tk.Button(btn_f, text="清", command=self.clear_sub_list, bg="#FFCCBC", font=("微软雅黑", 8, "bold"), width=4).pack(side=tk.RIGHT, padx=1)
        self.listbox = tk.Listbox(self.sub_win, font=("Consolas", 9), bd=1)
        self.listbox.pack(fill=tk.BOTH, expand=True)
        self.run_btn = tk.Button(self.sub_win, text="2. 执行录入", bg="#C8E6C9", font=("微软雅黑", 9, "bold"), command=self.confirm_and_run)
        self.run_btn.pack(fill=tk.X)
        self.update_sub_ui()

    def load_from_clipboard(self):
        try:
            raw = self.root.clipboard_get()
            self.current_sns = sorted(list(set([s.strip() for s in str(raw).split('\n') if s.strip()])))
            self.update_sub_ui()
        except: messagebox.showwarning("提示", "剪贴板为空")

    def clear_sub_list(self): self.current_sns = []; self.update_sub_ui()

    def update_sub_ui(self):
        if not hasattr(self, 'listbox'): return
        self.listbox.delete(0, tk.END)
        for sn in self.current_sns: self.listbox.insert(tk.END, sn)
        self.run_btn.config(text=f"2. 执行录入 ({len(self.current_sns)}条)" if self.current_sns else "2. 执行录入 (空)", state=tk.NORMAL if self.current_sns else tk.DISABLED)

    def confirm_and_run(self):
        if not self.current_sns: return
        sns_to_send = list(self.current_sns)
        self.sub_win.destroy()
        self.root.attributes("-alpha", self.work_alpha)
        threading.Thread(target=self._auto_core, args=(sns_to_send,), daemon=True).start()

    def _auto_core(self, sns):
        time.sleep(3) 
        e_del, m_del = float(self.spin_e1.get()), float(self.spin_mid.get())
        for sn in sns:
            kb_controller.type(sn); time.sleep(e_del)
            kb_controller.press(Key.enter); kb_controller.release(Key.enter)
            if self.use_double_enter.get(): 
                time.sleep(0.05)
                kb_controller.press(Key.enter); kb_controller.release(Key.enter)
            self.root.after(0, lambda s=sn: self.handle_scan(s, is_batch=True))
            time.sleep(m_del)
        self.root.after(0, lambda: self.root.attributes("-alpha", self.normal_alpha))

if __name__ == "__main__":
    root = tk.Tk(); app = UltimateMiniGuard(root); root.mainloop()
