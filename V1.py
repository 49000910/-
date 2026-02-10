import time
import threading
import winsound
import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk
import ctypes
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
            for line in f:
                if line.strip():
                    BARCODE_HISTORY.add(line.strip())
    except:
        pass

class UltimateMiniGuard:
    def __init__(self, root):
        self.root = root
        self.root.title("智能助手 v4.2 - 增强视觉版")
        self.root.geometry("450x240") 
        self.root.attributes("-topmost", True)
        
        # 颜色定义
        self.clr_default_bg = "#f0f0f0"
        self.clr_head_normal = "#90ee90" 
        self.clr_ok_green = "#2ecc71"      # 成功绿
        self.clr_ok_text_bg = "#d5f5e3"    # 日志框浅绿
        self.clr_dup_red = "#ff0000"       # 拦截红
        self.clr_dup_yellow = "#ffff00"    # 警示黄
        self.clr_dup_text_bg = "#ffb2b2"   # 日志框浅红
        
        self.root.configure(bg=self.clr_default_bg)

        # --- 参数区域 ---
        self.params_f = tk.Frame(self.root, bg=self.clr_head_normal, pady=3)
        self.params_f.pack(fill=tk.X)
        
        spin_opt = {"font": ("Consolas", 9), "width": 3, "from_": 0.0, "to": 9.9, "increment": 0.1}
        
        # PB 按钮：勾选时开启物理拦截（跳格全选），不勾选时仅闪灯报警
        self.use_pb = tk.BooleanVar(value=True)
        tk.Checkbutton(self.params_f, text="PB", variable=self.use_pb, bg=self.clr_head_normal).pack(side=tk.LEFT)
        
        tk.Label(self.params_f, text="E1:", bg=self.clr_head_normal).pack(side=tk.LEFT)
        self.spin_e1 = tk.Spinbox(self.params_f, **spin_opt)
        self.spin_e1.delete(0, "end"); self.spin_e1.insert(0, "0.1"); self.spin_e1.pack(side=tk.LEFT)

        self.use_double_enter = tk.BooleanVar(value=False)
        tk.Checkbutton(self.params_f, text="回2", variable=self.use_double_enter, bg=self.clr_head_normal).pack(side=tk.LEFT, padx=5)
        
        tk.Label(self.params_f, text="中:", bg=self.clr_head_normal).pack(side=tk.LEFT)
        self.spin_mid = tk.Spinbox(self.params_f, **spin_opt)
        self.spin_mid.delete(0, "end"); self.spin_mid.insert(0, "0.1"); self.spin_mid.pack(side=tk.LEFT)
        
        tk.Label(self.params_f, text="E2:", bg=self.clr_head_normal).pack(side=tk.LEFT)
        self.spin_e2 = tk.Spinbox(self.params_f, **spin_opt)
        self.spin_e2.delete(0, "end"); self.spin_e2.insert(0, "0.3"); self.spin_e2.pack(side=tk.LEFT)

        tk.Button(self.params_f, text="批量", command=self.pop_preview_window, bg="#e1e1e1", font=("微软雅黑", 8)).pack(side=tk.RIGHT, padx=5)

        # --- 日志区域 ---
        self.log_text = tk.Text(self.root, font=("Consolas", 9), bg="#ffffff", height=10, bd=1)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self.log_text.tag_config("dup", background="#ff0000", foreground="#ffffff") 
        self.log_text.tag_config("auto", foreground="#0000ff") 

        self.info_lbl = tk.Label(self.root, text=f"Total: {len(BARCODE_HISTORY)}", font=("微软雅黑", 8), bg=self.clr_default_bg)
        self.info_lbl.pack(side=tk.RIGHT, padx=5)

        self.listener = keyboard.Listener(on_press=self.on_press)
        self.listener.start()

    def update_ui_status(self, status):
        """更新UI：ok=绿，dup=红黄闪烁并留红"""
        if status == "ok":
            self.root.configure(bg=self.clr_ok_green)
            self.params_f.configure(bg=self.clr_ok_green)
            self.log_text.configure(bg=self.clr_ok_text_bg)
            self.info_lbl.configure(bg=self.clr_ok_green)
        elif status == "dup":
            def to_yellow(): self.root.configure(bg=self.clr_dup_yellow)
            def to_red(): 
                self.root.configure(bg=self.clr_dup_red)
                self.params_f.configure(bg=self.clr_dup_red)
                self.log_text.configure(bg=self.clr_dup_text_bg)
                self.info_lbl.configure(bg=self.clr_dup_red)
            
            to_red()
            self.root.after(150, to_yellow)
            self.root.after(300, to_red)
            self.root.after(450, to_yellow)
            self.root.after(600, to_red)

    def on_press(self, key):
        global LAST_KEY_TIME, SCAN_BUFFER
        now = time.time()
        interval = now - LAST_KEY_TIME
        LAST_KEY_TIME = now
        try:
            if hasattr(key, 'char') and key.char: char = key.char
            elif key == Key.enter: char = '\n'
            else: return

            if interval < SCAN_SPEED_THRESHOLD:
                if char == '\n':
                    barcode = "".join(SCAN_BUFFER).strip()
                    SCAN_BUFFER = []
                    if barcode: self.root.after(0, self.handle_scan, barcode)
                else: SCAN_BUFFER.append(char)
            else:
                SCAN_BUFFER = [char] if char != '\n' else []
        except: pass

    def handle_scan(self, barcode):
        # 1. 查重逻辑：只要重复，就执行变色闪烁和报警
        if barcode in BARCODE_HISTORY:
            winsound.Beep(1000, 800)
            self.update_ui_status("dup") # 执行红黄闪烁并保持红色
            self.log_text.insert("1.0", f"[发现重复] {barcode}\n", "dup")
            
            # 2. 物理拦截逻辑：只有 PB 勾选时才执行跳格全选
            if self.use_pb.get():
                kb_controller.press(Key.up)
                kb_controller.release(Key.up)
                time.sleep(0.05) 
                with kb_controller.pressed(Key.ctrl):
                    kb_controller.press('a'); kb_controller.release('a')
        else:
            # 正常条码：变绿色并记录
            self.update_ui_status("ok")
            BARCODE_HISTORY.add(barcode)
            with open(HISTORY_FILE, "a", encoding="utf-8") as f:
                f.write(barcode + "\n")
            self.log_text.insert("1.0", f"[采集成功] {barcode}\n")
            self.info_lbl.config(text=f"Total: {len(BARCODE_HISTORY)}")

    # --- 其他批量功能保持不变 ---
    def pop_preview_window(self):
        try: raw = self.root.clipboard_get()
        except: return
        lines = sorted(list(set([s.strip() for s in str(raw).split('\n') if s.strip()])))
        if not lines: return
        self.pv = tk.Toplevel(self.root); self.pv.title(f"批量预览"); self.pv.attributes("-topmost", True)
        self.tree = ttk.Treeview(self.pv, columns=("check", "barcode"), show="headings")
        self.tree.heading("check", text="状态"); self.tree.column("check", width=40)
        self.tree.heading("barcode", text="内容"); self.tree.column("barcode", width=300); self.tree.pack(fill=tk.BOTH, expand=True)
        for s in lines: self.tree.insert("", tk.END, values=("☐", s))
        self.tree.bind("<ButtonRelease-1>", lambda e: self.on_tree_click(e))
        tk.Button(self.pv, text="启动自动录入", bg="#4caf50", fg="white", command=self.run_auto).pack(fill=tk.X)

    def on_tree_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            v = list(self.tree.item(item, "values"))
            self.tree.item(item, values=("☑" if v == "☐" else "☐", v))

    def run_auto(self):
        to_run = [self.tree.item(i, "values") for i in self.tree.get_children() if self.tree.item(i, "values") == "☐"]
        if to_run: self.pv.destroy(); threading.Thread(target=self._auto_core, args=(to_run,), daemon=True).start()

    def _auto_core(self, sns):
        time.sleep(1.5)
        try:
            e1, mid, e2 = float(self.spin_e1.get()), float(self.spin_mid.get()), float(self.spin_e2.get())
            for sn in sns:
                with kb_controller.pressed(Key.ctrl): kb_controller.press('a'); kb_controller.release('a')
                time.sleep(0.05)
                self.root.clipboard_clear(); self.root.clipboard_append(sn); self.root.update()
                time.sleep(e1)
                with kb_controller.pressed(Key.ctrl): kb_controller.press('v'); kb_controller.release('v')
                time.sleep(mid)
                kb_controller.press(Key.enter); kb_controller.release(Key.enter)
                if self.use_double_enter.get():
                    time.sleep(0.1); kb_controller.press(Key.enter); kb_controller.release(Key.enter)
                self.root.after(0, lambda s=sn: self.update_auto_ui(s))
                time.sleep(e2)
        except: pass

    def update_auto_ui(self, sn):
        self.log_text.insert("1.0", f"[自动过站] {sn}\n", "auto")
        self.info_lbl.config(text=f"Total: {len(BARCODE_HISTORY)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = UltimateMiniGuard(root)
    root.mainloop()
