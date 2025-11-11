# STDF_FLASH.py
# 需求：已安裝並可 import 的 rust_stdf_helper 套件（含 stdf_to_log_sheet_stats_v6）
# 執行：python STDF_FLASH.py

import os
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --- Rust 綁定 ---
try:
    import cy_stdf_helper as rs
except Exception as e:
    raise SystemExit(
        "無法 import rust_stdf_helper。\n"
        "請先建置/安裝對應的 Python 擴充套件（maturin 或 wheel）。\n"
        f"原始錯誤：\n{e}"
    )

# ---- 提供給 Rust 的 progress_signal 與 stop_flag 物件 ----
class ProgressSignal:
    """Rust 端會呼叫 .emit(int)。設計 0..10000 對應到進度條 0..100。"""
    def __init__(self, tk_var: tk.DoubleVar):
        self._var = tk_var

    def emit(self, v: int):
        # 安全轉主執行緒更新
        def _update():
            pct = max(0.0, min(100.0, (v or 0) / 100.0))
            self._var.set(pct)
        try:
            # 若 tk 根視窗尚未建立 mainloop，直接設值也可；這裡保守用 after
            root.after(0, _update)
        except Exception:
            _update()


class StopFlag:
    """Rust 端會去取屬性 stop(bool)。這裡提供 thread-safe 的旗標。"""
    def __init__(self):
        self.stop = False


# ---- 主要 UI ----
class STDF_FLASH_UI(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.pack(fill="both", expand=True, padx=12, pady=12)

        # 狀態
        self.input_path = tk.StringVar(value="")
        self.format_var = tk.StringVar(value="xlsx")   # 預設 xlsx
        self.progress_var = tk.DoubleVar(value=0.0)
        self.running = False
        self.worker_thread = None
        self.stop_flag = StopFlag()

        # 標題
        title = ttk.Label(self, text="STDF_FLASH", font=("Segoe UI", 16, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w")

        # 選檔
        ttk.Label(self, text="STDF 檔案：").grid(row=1, column=0, sticky="e", pady=(12, 6))
        entry = ttk.Entry(self, textvariable=self.input_path, width=60)
        entry.grid(row=1, column=1, sticky="we", pady=(12, 6))
        btn_browse = ttk.Button(self, text="選擇檔案", command=self.browse_file)
        btn_browse.grid(row=1, column=2, padx=(8,0), pady=(12, 6))

        # 格式
        ttk.Label(self, text="輸出格式：").grid(row=2, column=0, sticky="e")
        fmt_box = ttk.Frame(self)
        fmt_box.grid(row=2, column=1, sticky="w")
        ttk.Radiobutton(fmt_box, text="Excel (.xlsx)", value="xlsx", variable=self.format_var).pack(side="left", padx=(0,12))
        ttk.Radiobutton(fmt_box, text="CSV (.csv)", value="csv", variable=self.format_var).pack(side="left")

        # 進度列
        ttk.Label(self, text="進度：").grid(row=3, column=0, sticky="e", pady=(12, 6))
        self.progress = ttk.Progressbar(self, variable=self.progress_var, maximum=100)
        self.progress.grid(row=3, column=1, columnspan=2, sticky="we", pady=(12, 6))

        # 控制鈕
        btn_box = ttk.Frame(self)
        btn_box.grid(row=4, column=0, columnspan=3, sticky="e", pady=(8, 0))
        self.btn_start = ttk.Button(btn_box, text="開始", command=self.on_start)
        self.btn_start.pack(side="left", padx=(0,8))
        self.btn_cancel = ttk.Button(btn_box, text="取消", command=self.on_cancel, state="disabled")
        self.btn_cancel.pack(side="left")

        # 拉伸
        self.columnconfigure(1, weight=1)

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="選擇 STDF 檔案",
            filetypes=[
                ("STDF or GZip", "*.stdf *.std *.stdz *.stdf.gz *.gz"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.input_path.set(path)

    def on_start(self):
        if self.running:
            return
        in_path = self.input_path.get().strip()
        if not in_path:
            messagebox.showwarning("提醒", "請先選擇 STDF 檔案")
            return
        if not os.path.isfile(in_path):
            messagebox.showerror("錯誤", "檔案不存在")
            return

        # 決定輸出路徑
        base = os.path.basename(in_path)
        # 去除一層 .gz（若有）
        if base.lower().endswith(".gz"):
            base_no_gz = os.path.splitext(base)[0]
        else:
            base_no_gz = base

        name_no_ext, _ext = os.path.splitext(base_no_gz)
        out_fmt = self.format_var.get()
        out_ext = ".xlsx" if out_fmt == "xlsx" else ".csv"
        out_path = os.path.join(os.path.dirname(in_path), f"{name_no_ext}{out_ext}")

        # 建立 progress 與 stop flag
        self.progress_var.set(0.0)
        self.stop_flag = StopFlag()
        progress_signal = ProgressSignal(self.progress_var)

        # 固定 TestNumberOnly
        test_id_type = rs.TestIDType.TestNumberOnly

        # UI 狀態
        self.running = True
        self.btn_start.config(state="disabled")
        self.btn_cancel.config(state="normal")

        # 執行緒轉檔
        def _work():
            try:
                # 呼叫 rust 函式
                rs.stdf_to_log_sheet_stats_v6(
                    in_path,
                    out_path,
                    test_id_type,
                    progress_signal,  # 需有 emit(int)
                    self.stop_flag,   # 需有屬性 stop: bool
                )
                if self.stop_flag.stop:
                    raise RuntimeError("已取消")
                # 成功
                root.after(0, lambda: messagebox.showinfo("完成", f"已輸出：\n{out_path}"))
            except Exception as e:
                tb = traceback.format_exc(limit=4)
                root.after(0, lambda: messagebox.showerror("失敗", f"{e}\n\n{tb}"))
            finally:
                root.after(0, self._reset_ui)

        self.worker_thread = threading.Thread(target=_work, daemon=True)
        self.worker_thread.start()

    def on_cancel(self):
        if not self.running:
            return
        self.stop_flag.stop = True
        self.btn_cancel.config(state="disabled")

    def _reset_ui(self):
        self.running = False
        self.btn_start.config(state="normal")
        self.btn_cancel.config(state="disabled")


if __name__ == "__main__":
    root = tk.Tk()
    root.title("STDF_FLASH")
    try:
        # Windows 預設好看一點
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    style = ttk.Style()
    try:
        style.theme_use("vista")
    except Exception:
        pass
    app = STDF_FLASH_UI(root)
    root.minsize(640, 220)
    root.mainloop()
