import time
import threading
import winsound
import os
import tkinter as tk
from tkinter import messagebox, ttk
from pynput import keyboard
from pynput.keyboard import Controller, Key

# --- 核心配置 ---
HISTORY_FILE = "barcode_history.txt"
BARCODE_HISTORY = set()
SCAN_BUFFER = []
LAST_KEY_TIME = 0
SCAN_SPEED_THRESHOLD = 0.05 # 拦截判定阈值：0.05秒
kb_controller = Controller()

if os.path.exists(HISTORY_FILE):
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            for line in f:
                if line.strip(): BARCODE_HISTORY.add(line.strip())
    except: pass

class UltimateMiniGuard:
    def __init__(self, root):
        self.root = root
        self.root.geometry("450x260") 
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", 0.92)
        self.root.overrideredirect(True) 

        # 舒服的配色方案
        self.clr_title_bar = "#80CBC4"    
        self.clr_head_normal = "#B2DFDB"  
        self.clr_default_bg = "#ECEFF1"   
        self.clr_ok_green = "#A5D6A7"     
        self.clr_dup_red = "#EF9A9A"      
        self.clr_dup_yellow = "#FFF59D"   
        
        self.root.configure(bg=self.clr_default_bg)

        # --- 1. 自定义标题栏 ---
        self.title_bar = tk.Frame(self.root, bg=self.clr_title_bar, height=25)
        self.title_bar.pack(fill=tk.X)
        self.title_bar.pack_propagate(False)

        self.title_lbl = tk.Label(self.title_bar, text=" 智能助手 v4.5 - 网页表单版", 
                                 fg="#004D40", bg=self.clr_title_bar, font=("微软雅黑", 9, "bold"))
        self.title_lbl.pack(side=tk.LEFT)

        tk.Button(self.title_bar, text="✕", command=root.quit, bg="#FF7043", fg="white", 
                  font=("Arial", 9, "bold"), bd=0, cursor="hand2", padx=10).pack(side=tk.RIGHT)

        for widget in [self.title_bar, self.title_lbl, self.root]:
            widget.bind("<Button-1>", self.start_move)
            widget.bind("<B1-Motion>", self.do_move)

        # --- 2. 参数区域 ---
        self.params_f = tk.Frame(self.root, bg=self.clr_head_normal, pady=3)
        self.params_f.pack(fill=tk.X)
        
        spin_opt = {"font": ("Consolas", 9), "width": 3, "from_": 0.0, "to": 9.9, "increment": 0.1}
        self.use_pb = tk.BooleanVar(value=True)
        tk.Checkbutton(self.params_f, text="PB", variable=self.use_pb, bg=self.clr_head_normal).pack(side=tk.LEFT, padx=2)
        
        tk.Label(self.params_f, text="E1:", bg=self.clr_head_normal).pack(side=tk.LEFT)
        self.spin_e1 = tk.Spinbox(self.params_f, **spin_opt)
        self.spin_e1.delete(0, "end"); self.spin_e1.insert(0, "0.1"); self.spin_e1.pack(side=tk.LEFT)

        tk.Label(self.params_f, text="中:", bg=self.clr_head_normal).pack(side=tk.LEFT)
        self.spin_mid = tk.Spinbox(self.params_f, **spin_opt)
        self.spin_mid.delete(0, "end"); self.spin_mid.insert(0, "0.8"); self.spin_mid.pack(side=tk.LEFT)
        
        tk.Button(self.params_f, text="批量", command=self.pop_preview_window, bg="#CFD8DC", font=("微软雅黑", 8), relief=tk.FLAT).pack(side=tk.RIGHT, padx=5)

        # --- 3. 日志区域 ---
        self.log_text = tk.Text(self.root, font=("Consolas", 9), bg="#ffffff", height=10, bd=0, padx=5, pady=5)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text.tag_config("dup", background="#FFEBEE", foreground="#C62828")

        self.info_lbl = tk.Label(self.root, text=f"Total: {len(BARCODE_HISTORY)}", font=("微软雅黑", 8), bg=self.clr_default_bg)
        self.info_lbl.pack(side=tk.RIGHT, padx=5)

        self.listener = keyboard.Listener(on_press=self.on_press)
        self.listener.start()

    def start_move(self, event): self.x, self.y = event.x, event.y
    def do_move(self, event):
        x = self.root.winfo_x() + (event.x - self.x)
        y = self.root.winfo_y() + (event.y - self.y)
        self.root.geometry(f"+{x}+{y}")

    def set_all_colors(self, color, text_bg=None):
        t_bg = text_bg if text_bg else color
        self.root.configure(bg=color); self.title_bar.configure(bg=color)
        self.title_lbl.configure(bg=color); self.params_f.configure(bg=color)
        self.log_text.configure(bg=t_bg); self.info_lbl.configure(bg=color)

    def flash_warning(self):
        """1.5秒警报闪烁"""
        end_time = time.time() + 1.5
        def do_flash():
            if time.time() < end_time:
                curr = self.root.cget("bg")
                new_c = self.clr_dup_yellow if curr == self.clr_dup_red else self.clr_dup_red
                self.set_all_colors(new_c)
                self.root.after(120, do_flash)
            else:
                self.set_all_colors(self.clr_dup_red, "#FFEBEE")
                self.title_bar.configure(bg="#D32F2F"); self.title_lbl.configure(bg="#D32F2F", fg="white")
        do_flash()

    def handle_scan(self, barcode):
        if barcode in BARCODE_HISTORY:
            winsound.Beep(1200, 400); self.flash_warning()
            self.log_text.insert("1.0", f"[拦截重复] {barcode}\n", "dup")
            if self.use_pb.get():
                # 网页拦截：回退并全选
                with kb_controller.pressed(Key.shift):
                    kb_controller.press(Key.tab); kb_controller.release(Key.tab)
                time.sleep(0.08) 
                with kb_controller.pressed(Key.ctrl):
                    kb_controller.press('a'); kb_controller.release('a')
        else:
            self.set_all_colors(self.clr_ok_green, "#E8F5E9")
            self.title_bar.configure(bg="#66BB6A"); self.title_lbl.configure(bg="#66BB6A", fg="#E8F5E9")
            BARCODE_HISTORY.add(barcode)
            with open(HISTORY_FILE, "a", encoding="utf-8") as f: f.write(barcode + "\n")
            self.log_text.insert("1.0", f"[扫描成功] {barcode}\n")
            self.info_lbl.config(text=f"Total: {len(BARCODE_HISTORY)}")

    def on_press(self, key):
        global LAST_KEY_TIME, SCAN_BUFFER
        now = time.time(); interval = now - LAST_KEY_TIME; LAST_KEY_TIME = now
        try:
            if hasattr(key, 'char') and key.char: char = key.char
            elif key == Key.enter: char = '\n'
            else: return
            if interval < SCAN_SPEED_THRESHOLD:
                if char == '\n':
                    bc = "".join(SCAN_BUFFER).strip(); SCAN_BUFFER = []
                    if bc: self.root.after(0, self.handle_scan, bc)
                else: SCAN_BUFFER.append(char)
            else: SCAN_BUFFER = [char] if char != '\n' else []
        except: pass

    def pop_preview_window(self):
        try: raw = self.root.clipboard_get()
        except: return
        lines = sorted(list(set([s.strip() for s in str(raw).split('\n') if s.strip()])))
        if not lines: return
        self.pv = tk.Toplevel(self.root); self.pv.title("预览"); self.pv.attributes("-alpha", 0.95); self.pv.attributes("-topmost", True)
        self.tree = ttk.Treeview(self.pv, columns=("check", "barcode"), show="headings")
        self.tree.heading("check", text="状态"); self.tree.column("check", width=40); self.tree.heading("barcode", text="内容"); self.tree.column("barcode", width=300); self.tree.pack(fill=tk.BOTH, expand=True)
        for s in lines: self.tree.insert("", tk.END, values=("☐", s))
        self.tree.bind("<ButtonRelease-1>", lambda e: self.on_tree_click(e))
        tk.Button(self.pv, text="🚀 开始批量录入", bg="#81C784", command=self.run_auto).pack(fill=tk.X)

    def on_tree_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            v = list(self.tree.item(item, "values"))
            self.tree.item(item, values=("☑" if v[0] == "☐" else "☐", v[1]))

    def run_auto(self):
        to_run = [self.tree.item(i, "values")[1] for i in self.tree.get_children() if self.tree.item(i, "values")[0] == "☐"]
        if to_run: self.pv.destroy(); threading.Thread(target=self._auto_core, args=(to_run,), daemon=True).start()

    def _auto_core(self, sns):
        time.sleep(1.5)
        try:
            e1, mid = float(self.spin_e1.get()), float(self.spin_mid.get())
            for sn in sns:
                with kb_controller.pressed(Key.ctrl):
                    kb_controller.press('a'); kb_controller.release('a')
                time.sleep(0.05); self.root.clipboard_clear(); self.root.clipboard_append(sn); self.root.update()
                time.sleep(e1); with kb_controller.pressed(Key.ctrl):
                    kb_controller.press('v'); kb_controller.release('v')
                time.sleep(mid); kb_controller.press(Key.enter); kb_controller.release(Key.enter)
                self.root.after(0, lambda s=sn: self.update_auto_ui(s)); time.sleep(0.3)
        except: pass

    def update_auto_ui(self, sn):
        self.log_text.insert("1.0", f"[批量完成] {sn}\n")
        self.info_lbl.config(text=f"Total: {len(BARCODE_HISTORY)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = UltimateMiniGuard(root)
    root.mainloop()
