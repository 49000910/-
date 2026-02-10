import time
import threading
import winsound
import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk
import ctypes
from pynput.keyboard import Controller, Key

# --- 核心配置 ---
HISTORY_FILE = "barcode_history.txt"
BARCODE_HISTORY = set()
kb_controller = Controller()

# 加载历史记录
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
        self.root.title("智能助手 v4.2")
        self.root.geometry("450x240") 
        self.root.attributes("-topmost", True)
        
        self.clr_default_bg = "#f0f0f0"
        self.clr_head_normal = "#90ee90" 
        self.clr_log_normal = "#ffffff" 
        self.root.configure(bg=self.clr_default_bg)

        # --- 顶部工具栏 ---
        params_f = tk.Frame(self.root, bg=self.clr_head_normal, pady=3)
        params_f.pack(fill=tk.X)
        
        spin_opt = {"font": ("Consolas", 9), "width": 3, "from_": 0.0, "to": 9.9, "increment": 0.1}
        self.use_pb = tk.BooleanVar(value=True)
        tk.Checkbutton(params_f, text="PB", variable=self.use_pb, bg=self.clr_head_normal).pack(side=tk.LEFT)
        
        tk.Label(params_f, text="E1:", bg=self.clr_head_normal).pack(side=tk.LEFT)
        self.spin_e1 = tk.Spinbox(params_f, **spin_opt)
        self.spin_e1.delete(0, "end")
        self.spin_e1.insert(0, "0.1")
        self.spin_e1.pack(side=tk.LEFT)

        self.use_double_enter = tk.BooleanVar(value=False)
        tk.Checkbutton(params_f, text="回2", variable=self.use_double_enter, bg=self.clr_head_normal).pack(side=tk.LEFT, padx=5)
        
        tk.Label(params_f, text="中:", bg=self.clr_head_normal).pack(side=tk.LEFT)
        self.spin_mid = tk.Spinbox(params_f, **spin_opt)
        self.spin_mid.delete(0, "end")
        self.spin_mid.insert(0, "0.1")
        self.spin_mid.pack(side=tk.LEFT)
        
        tk.Label(params_f, text="E2:", bg=self.clr_head_normal).pack(side=tk.LEFT)
        self.spin_e2 = tk.Spinbox(params_f, **spin_opt)
        self.spin_e2.delete(0, "end")
        self.spin_e2.insert(0, "0.3")
        self.spin_e2.pack(side=tk.LEFT)

        tk.Button(params_f, text="批量", command=self.pop_preview_window, bg="#e1e1e1", font=("微软雅黑", 8)).pack(side=tk.RIGHT, padx=5)

        # --- 日志显示 ---
        self.log_text = tk.Text(self.root, font=("Consolas", 9), bg=self.clr_log_normal, height=10, bd=1)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self.log_text.tag_config("auto", foreground="#0000ff") 

        self.info_lbl = tk.Label(self.root, text=f"Total: {len(BARCODE_HISTORY)}", font=("微软雅黑", 8), bg=self.clr_default_bg)
        self.info_lbl.pack(side=tk.RIGHT, padx=5)

    def get_clipboard_text(self):
        try:
            return self.root.clipboard_get()
        except:
            return ""

    def pop_preview_window(self):
        raw_text = self.get_clipboard_text()
        if not raw_text.strip():
            messagebox.showinfo("提示", "剪贴板内容为空")
            return
        
        lines = sorted(list(set([s.strip() for s in raw_text.splitlines() if s.strip()])))
        
        self.pv = tk.Toplevel(self.root)
        self.pv.title(f"预览: {len(lines)}条")
        self.pv.geometry("400x400")
        self.pv.attributes("-topmost", True)
        
        self.tree = ttk.Treeview(self.pv, columns=("check", "barcode"), show="headings")
        self.tree.heading("check", text="状态")
        self.tree.heading("barcode", text="条码内容")
        self.tree.column("check", width=50, anchor="center")
        self.tree.pack(fill=tk.BOTH, expand=True)

        for s in lines:
            self.tree.insert("", tk.END, values=("☐", s))

        self.tree.bind("<ButtonRelease-1>", self.on_tree_click)
        tk.Button(self.pv, text="🚀 开始批量录入", bg="#4caf50", fg="white", 
                  command=self.prepare_and_run).pack(fill=tk.X, pady=5)

    def on_tree_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            val = list(self.tree.item(item, "values"))
            val[0] = "☑" if val[0] == "☐" else "☐"
            self.tree.item(item, values=val)

    def prepare_and_run(self):
        # 提取未打叉的项目
        to_run = [self.tree.item(i, "values")[1] for i in self.tree.get_children() if self.tree.item(i, "values")[0] == "☐"]
        if to_run:
            self.pv.destroy()
            threading.Thread(target=self._auto_run, args=(to_run,), daemon=True).start()

    def _auto_run(self, sns):
        """核心自动化逻辑"""
        time.sleep(2.0) # 给用户时间切换窗口
        try:
            e1 = float(self.spin_e1.get())
            mid = float(self.spin_mid.get())
            e2 = float(self.spin_e2.get())
            
            for sn in sns:
                # 1. 全选旧内容
                with kb_controller.pressed(Key.ctrl):
                    kb_controller.press('a')
                    kb_controller.release('a')
                time.sleep(0.05)
                kb_controller.press(Key.backspace)
                kb_controller.release(Key.backspace)

                # 2. 写入剪贴板并粘贴
                self.root.clipboard_clear()
                self.root.clipboard_append(sn)
                time.sleep(e1)
                
                with kb_controller.pressed(Key.ctrl):
                    kb_controller.press('v')
                    kb_controller.release('v')
                
                time.sleep(mid)
                
                # 3. 回车确认
                kb_controller.press(Key.enter)
                kb_controller.release(Key.enter)
                
                if self.use_double_enter.get():
                    time.sleep(0.1)
                    kb_controller.press(Key.enter)
                    kb_controller.release(Key.enter)

                # 4. 更新UI日志
                self.root.after(0, self.update_log, sn)
                time.sleep(e2)
                
            winsound.Beep(1000, 500)
            self.root.after(0, lambda: messagebox.showinfo("完成", "批量任务执行完毕"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"运行中断: {e}"))

    def update_log(self, sn):
        self.log_text.insert(tk.END, f"已录入: {sn}\n", "auto")
        self.log_text.see(tk.END)
        if sn not in BARCODE_HISTORY:
            BARCODE_HISTORY.add(sn)
            with open(HISTORY_FILE, "a", encoding="utf-8") as f:
                f.write(sn + "\n")
        self.info_lbl.config(text=f"Total: {len(BARCODE_HISTORY)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = UltimateMiniGuard(root)
    root.mainloop()
