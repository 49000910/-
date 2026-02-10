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

class UltimateMiniGuard:
    def __init__(self, root):
        self.root = root
        self.root.title("智能助手 v4.0")
        self.root.geometry("450x240") 
        self.root.attributes("-topmost", True)
        
        self.clr_default_bg = "#f0f0f0"; self.clr_head_normal = "#90ee90" 
        self.clr_ok_static = "#2ecc71"; self.clr_dup_red = "#ff0000"; self.clr_dup_yellow = "#ffff00"
        self.clr_log_normal = "#ffffff" 
        
        self.flash_timer = None; self.is_flashing = False 
        self.current_state_color = self.clr_default_bg
        self.root.configure(bg=self.clr_default_bg)

        # 1. 顶部控制区
        self.head_f = tk.Frame(self.root, bg=self.clr_head_normal, pady=3); self.head_f.pack(fill=tk.X)
        params_f = tk.Frame(self.head_f, bg=self.clr_head_normal); params_f.pack(fill=tk.X, padx=5)
        
        spin_opt = {"font": ("Consolas", 9), "width": 3, "from_": 0.0, "to": 9.9, "increment": 0.1}
        
        self.use_pb = tk.BooleanVar(value=True)
        tk.Checkbutton(params_f, text="PB", variable=self.use_pb, bg=self.clr_head_normal, font=("Arial", 8, "bold")).pack(side=tk.LEFT)
        
        tk.Label(params_f, text="E1:", bg=self.clr_head_normal).pack(side=tk.LEFT)
        self.spin_e1 = tk.Spinbox(params_f, **spin_opt)
        self.spin_e1.delete(0, "end"); self.spin_e1.insert(0, "0.1"); self.spin_e1.pack(side=tk.LEFT)

        self.use_double_enter = tk.BooleanVar(value=False)
        tk.Checkbutton(params_f, text="回2", variable=self.use_double_enter, bg=self.clr_head_normal).pack(side=tk.LEFT, padx=(5,0))
        
        tk.Label(params_f, text="中:", bg=self.clr_head_normal).pack(side=tk.LEFT)
        self.spin_mid = tk.Spinbox(params_f, **spin_opt)
        self.spin_mid.delete(0, "end"); self.spin_mid.insert(0, "0.1"); self.spin_mid.pack(side=tk.LEFT)
        
        tk.Label(params_f, text="E2:", bg=self.clr_head_normal).pack(side=tk.LEFT, padx=(5,0))
        self.spin_e2 = tk.Spinbox(params_f, **spin_opt)
        self.spin_e2.delete(0, "end"); self.spin_e2.insert(0, "0.3"); self.spin_e2.pack(side=tk.LEFT)

        tk.Button(params_f, text="批量", command=self.pop_preview_window, bg="#e1e1e1", font=("微软雅黑", 8)).pack(side=tk.RIGHT, padx=2)

        # 2. 日志显示区
        self.log_text = tk.Text(self.root, font=("Consolas", 9), bg=self.clr_log_normal, height=10, bd=1)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self.log_text.tag_config("dup", background="#ffb2b2", foreground="#b22222")

        # 3. 底部状态栏
        self.info_f = tk.Frame(self.root, bg=self.clr_default_bg); self.info_f.pack(fill=tk.X)
        self.info_lbl = tk.Label(self.info_f, text=f"Total: {len(BARCODE_HISTORY)}", font=("微软雅黑", 8), bg=self.clr_default_bg)
        self.info_lbl.pack(side=tk.RIGHT, padx=5); self.log_count = len(BARCODE_HISTORY)

    def _apply_ui_color(self, main_clr, log_clr):
        self.root.configure(bg=main_clr); self.head_f.configure(bg=main_clr)
        self.log_text.configure(bg=log_clr); self.info_f.configure(bg=main_clr); self.info_lbl.configure(bg=main_clr)

    def trigger_alarm(self, is_dup):
        if self.flash_timer: self.root.after_cancel(self.flash_timer)
        self.is_flashing = False
        if not is_dup:
            # 正常扫码：变绿常驻，不全选
            winsound.Beep(800, 150); self.current_state_color = self.clr_ok_static
            self._apply_ui_color(self.clr_ok_static, "#d5f5e3")
        else:
            # 重复扫码：红黄闪烁
            winsound.Beep(1200, 600); self.is_flashing = True; self._dup_flash_step(0)
            self.flash_timer = self.root.after(1500, self.stop_flash)

    def _dup_flash_step(self, count):
        if not self.is_flashing: return
        clr = self.clr_dup_red if count % 2 == 0 else self.clr_dup_yellow
        self._apply_ui_color(clr, clr); self.flash_timer = self.root.after(100, lambda: self._dup_flash_step(count + 1))

    def stop_flash(self):
        self.is_flashing = False
        target_log = "#d5f5e3" if self.current_state_color == self.clr_ok_static else self.clr_log_normal
        self._apply_ui_color(self.current_state_color, target_log)

    def add_log(self, code, status_tag=None):
        self.log_text.config(state=tk.NORMAL); ts = time.strftime("%H:%M:%S"); self.log_count += 1
        self.log_text.insert(tk.END, f"[{self.log_count:02d}] {ts} {code}\n", status_tag)
        self.log_text.see(tk.END); self.log_text.config(state=tk.DISABLED)
        self.info_lbl.config(text=f"Total: {len(BARCODE_HISTORY)}")

    def update_monitor(self, code, is_dup):
        self.trigger_alarm(is_dup)
        if is_dup:
            self.add_log(code, "dup")
            # 扫码拦截报错时的PB逻辑（不改）
            if self.use_pb.get():
                with kb_controller.pressed(Key.shift): kb_controller.press(Key.tab); kb_controller.release(Key.tab)
                time.sleep(0.2); with kb_controller.pressed(Key.ctrl): kb_controller.press('a'); kb_controller.release('a')
        else: self.add_log(code)

    def pop_preview_window(self):
        try: sns = sorted(list(set(s.strip() for s in self.root.clipboard_get().split('\n') if s.strip())))
        except: return
        pv = tk.Toplevel(self.root); pv.title("核对清单"); pv.geometry("240x300"); pv.attributes("-topmost", True)
        lb = tk.Listbox(pv); lb.pack(fill=tk.BOTH, expand=True)
        for s in sns: lb.insert(tk.END, s)
        tk.Button(pv, text="启动录入", bg="#4caf50", fg="white", command=lambda: [self.execute_auto(list(lb.get(0, tk.END))), pv.destroy()]).pack(fill=tk.X)

    def execute_auto(self, sns):
        threading.Thread(target=self._auto_run, args=(sns,), daemon=True).start()

    def _auto_run(self, sns):
        """批量录入逻辑：使用 0.1s 稳健延迟"""
        time.sleep(1.5) 
        try:
            e1 = float(self.spin_e1.get()); mid_delay = float(self.spin_mid.get()); e2 = float(self.spin_e2.get())
            for sn in sns:
                # 1. 全选 (批量模式特有)
                with kb_controller.pressed(Key.ctrl): kb_controller.press('a'); kb_controller.release('a')
                time.sleep(0.1) 
                # 2. 写入剪贴板
                self.root.after(0, lambda x=sn: [self.root.clipboard_clear(), self.root.clipboard_append(x)])
                time.sleep(e1) 
                # 3. 粘贴
                with kb_controller.pressed(Key.ctrl): kb_controller.press('v'); kb_controller.release('v')
                time.sleep(0.1) 
                # 4. 回车
                kb_controller.press(Key.enter); kb_controller.release(Key.enter)
                if self.use_double_enter.get():
                    time.sleep(mid_delay) 
                    kb_controller.press(Key.enter); kb_controller.release(Key.enter)
                time.sleep(e2)
                
                if sn not in BARCODE_HISTORY:
                    BARCODE_HISTORY.add(sn); 
                    with open(HISTORY_FILE, "a", encoding="utf-8") as f: f.write(f"{sn}\n")
                self.root.after(0, self.add_log, sn, "auto")
        except: pass

def on_press(key):
    global LAST_KEY_TIME, SCAN_BUFFER
    now = time.time()
    try:
        char = key.char if hasattr(key, 'char') else ('\n' if key == Key.enter else None)
        if not char: return
        if now - LAST_KEY_TIME < SCAN_SPEED_THRESHOLD:
            if char == '\n':
                barcode = "".join(SCAN_BUFFER).strip()
                if barcode:
                    is_dup = barcode in BARCODE_HISTORY
                    if not is_dup:
                        BARCODE_HISTORY.add(barcode)
                        with open(HISTORY_FILE, "a", encoding="utf-8") as f: f.write(f"{barcode}\n")
                    app.root.after(0, app.update_monitor, barcode, is_dup)
                SCAN_BUFFER = []
            else: SCAN_BUFFER.append(char)
        else: SCAN_BUFFER = [char] if char != '\n' else []
        LAST_KEY_TIME = now
    except: pass

if __name__ == "__main__":
    root = tk.Tk(); app = UltimateMiniGuard(root)
    listener = keyboard.Listener(on_press=on_press); listener.start(); root.mainloop()
