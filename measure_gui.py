#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CATIA 零件测量与提取 — 图形界面
用复选框选择要提取的内容(几何测量/用户参数/三维标注/整机截图/标注截图)
以及 Excel 输出哪些页面, 后台调用 measure_catia 完成测量与导出。

用法:
  python measure_catia.py --gui    # 或直接运行本文件
  python measure_gui.py
"""
import json
import os
import queue
import sys
import threading

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import measure_catia as mc

# Excel 可输出页面(与 write_excel 的 sheets 参数一致)
SHEET_NAMES = ["测量结果", "用户参数", "参数汇总", "三维标注", "MBD参数", "MBD捕获截图"]


class App:
    def __init__(self, root):
        self.root = root
        root.title("CATIA 零件测量与提取")
        root.geometry("760x740")
        self.q = queue.Queue()
        self.worker_running = False

        # ---- 测量设置 ----
        frm0 = ttk.LabelFrame(root, text="测量设置")
        frm0.pack(fill="x", padx=8, pady=6)
        ttk.Label(frm0, text="零件目录:").grid(row=0, column=0, padx=6, pady=4, sticky="w")
        self.dir_var = tk.StringVar(value="test")
        ttk.Entry(frm0, textvariable=self.dir_var, width=50).grid(row=0, column=1, padx=4, pady=4, sticky="we")
        ttk.Button(frm0, text="浏览…", command=self.browse).grid(row=0, column=2, padx=4, pady=4)
        ttk.Label(frm0, text="密度 (g/cm³):").grid(row=1, column=0, padx=6, pady=4, sticky="w")
        self.density_var = tk.StringVar(value=str(mc.DEFAULT_DENSITY))
        ttk.Entry(frm0, textvariable=self.density_var, width=12).grid(row=1, column=1, padx=4, pady=4, sticky="w")
        frm0.columnconfigure(1, weight=1)

        # ---- 提取内容 ----
        frm1 = ttk.LabelFrame(root, text="提取内容(可多选)")
        frm1.pack(fill="x", padx=8, pady=6)
        self.cb_measure = tk.BooleanVar(value=True)
        self.cb_params = tk.BooleanVar(value=True)
        self.cb_anns = tk.BooleanVar(value=True)
        self.cb_part_shot = tk.BooleanVar(value=True)
        self.cb_ann_shot = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm1, text="几何测量(体积/表面积/质量/重量/包络体)", variable=self.cb_measure) \
            .grid(row=0, column=0, sticky="w", padx=10, pady=3)
        ttk.Checkbutton(frm1, text="用户自定义参数", variable=self.cb_params) \
            .grid(row=1, column=0, sticky="w", padx=10, pady=3)
        ttk.Checkbutton(frm1, text="三维标注(收集标注文本)", variable=self.cb_anns) \
            .grid(row=2, column=0, sticky="w", padx=10, pady=3)
        ttk.Checkbutton(frm1, text="零件整机截图(600×500, 写入测量结果截图列)", variable=self.cb_part_shot) \
            .grid(row=3, column=0, sticky="w", padx=10, pady=3)
        ttk.Checkbutton(frm1, text="三维标注逐个截图(较耗时)", variable=self.cb_ann_shot, command=self.on_ann_shot) \
            .grid(row=4, column=0, sticky="w", padx=10, pady=3)

        # ---- MBD 提取(源自 MBD.txt VBA 移植, 勾选后按选项提取, 结果写入同一个 Excel) ----
        frm4 = ttk.LabelFrame(root, text="MBD 提取(写入同一个 Excel, 可多选)")
        frm4.pack(fill="x", padx=8, pady=6)
        self.cb_mbd_params = tk.BooleanVar(value=False)
        self.cb_mbd_caps = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm4, text="MBD 指定参数(Part Notes:/Standard Notes:/ECCN 等 -> Excel「MBD参数」页)",
                        variable=self.cb_mbd_params).grid(row=0, column=0, sticky="w", padx=10, pady=3)
        ttk.Checkbutton(frm4, text="MBD Capture 捕获视图逐个截图(较耗时 -> Excel「MBD捕获截图」页)",
                        variable=self.cb_mbd_caps).grid(row=1, column=0, sticky="w", padx=10, pady=3)

        # ---- Excel 输出页面 ----
        frm2 = ttk.LabelFrame(root, text="Excel 输出页面(可多选)")
        frm2.pack(fill="x", padx=8, pady=6)
        self.s_measure = tk.BooleanVar(value=True)
        self.s_params = tk.BooleanVar(value=True)
        self.s_ps = tk.BooleanVar(value=True)
        self.s_anns = tk.BooleanVar(value=True)
        self.s_mbd_params = tk.BooleanVar(value=True)
        self.s_mbd_caps = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm2, text="测量结果(含截图列)", variable=self.s_measure) \
            .grid(row=0, column=0, sticky="w", padx=10, pady=2)
        ttk.Checkbutton(frm2, text="用户参数", variable=self.s_params) \
            .grid(row=0, column=1, sticky="w", padx=10, pady=2)
        ttk.Checkbutton(frm2, text="参数汇总", variable=self.s_ps) \
            .grid(row=0, column=2, sticky="w", padx=10, pady=2)
        ttk.Checkbutton(frm2, text="三维标注(单页, 含截图)", variable=self.s_anns) \
            .grid(row=1, column=0, sticky="w", padx=10, pady=2)
        ttk.Checkbutton(frm2, text="MBD参数(含几何图形集参数, 去重)", variable=self.s_mbd_params) \
            .grid(row=1, column=1, sticky="w", padx=10, pady=2)
        ttk.Checkbutton(frm2, text="MBD捕获截图(按件号分页, 含图片)", variable=self.s_mbd_caps) \
            .grid(row=2, column=0, sticky="w", padx=10, pady=2)

        # ---- 按钮与状态 ----
        btns = ttk.Frame(root)
        btns.pack(fill="x", padx=8, pady=4)
        self.start_btn = ttk.Button(btns, text="开始提取", command=self.start)
        self.start_btn.pack(side="left", padx=4)
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(btns, textvariable=self.status_var).pack(side="left", padx=12)

        # ---- 日志区 ----
        frm3 = ttk.LabelFrame(root, text="日志")
        frm3.pack(fill="both", expand=True, padx=8, pady=6)
        self.log_text = tk.Text(frm3, height=14, wrap="word")
        sb = ttk.Scrollbar(frm3, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="both", expand=True)

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after_id = root.after(100, self.poll)

    # ---------- 界面交互 ----------
    def browse(self):
        d = filedialog.askdirectory(title="选择零件目录")
        if d:
            self.dir_var.set(d)

    def on_ann_shot(self):
        # 标注逐个截图依赖标注收集, 勾选时自动补选
        if self.cb_ann_shot.get():
            self.cb_anns.set(True)

    def log_append(self, s):
        self.log_text.insert("end", s)
        self.log_text.see("end")

    def poll(self):
        """主线程轮询后台队列, 刷新日志与状态"""
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self.log_append(payload)
                elif kind == "done":
                    self.worker_running = False
                    self.start_btn.config(state="normal")
                    self.status_var.set("完成" if payload and not payload.startswith("出错") else "出错")
                    self.log_append(payload + "\n")
                    messagebox.showinfo("完成", payload)
        except queue.Empty:
            pass
        self.after_id = self.root.after(100, self.poll)

    def on_close(self):
        if self.worker_running:
            messagebox.showinfo("提示", "正在处理中, 请稍候再关闭窗口。")
            return
        try:
            self.root.after_cancel(self.after_id)
        except Exception:
            pass
        self.root.destroy()

    def start(self):
        if self.worker_running:
            return
        directory = os.path.abspath(self.dir_var.get().strip())
        if not os.path.isdir(directory):
            messagebox.showerror("错误", "目录不存在:\n" + directory)
            return
        try:
            density = float(self.density_var.get().strip())
        except ValueError:
            messagebox.showerror("错误", "密度必须为数字(如 7.85)")
            return
        if density <= 0:
            messagebox.showerror("错误", "密度必须大于 0")
            return
        sheets = [n for n in SHEET_NAMES if {
            "测量结果": self.s_measure, "用户参数": self.s_params,
            "参数汇总": self.s_ps, "三维标注": self.s_anns,
            "MBD参数": self.s_mbd_params, "MBD捕获截图": self.s_mbd_caps,
        }[n].get()]
        if not sheets:
            messagebox.showerror("错误", "请至少勾选一个 Excel 输出页面")
            return
        opts = {
            "measure": self.cb_measure.get(),
            "user_params": self.cb_params.get(),
            "annotations": self.cb_anns.get(),
            "part_shot": self.cb_part_shot.get(),
            "mbd_params": self.cb_mbd_params.get(),
            "mbd_captures": self.cb_mbd_caps.get(),
        }
        do_ann_shots = self.cb_ann_shot.get()

        self.worker_running = True
        self.start_btn.config(state="disabled")
        self.status_var.set("处理中…")
        self.log_text.delete("1.0", "end")

        # 把模块日志转发到界面队列(后台线程调用 log 时入队, 主线程 poll 刷新)
        mc.log = lambda msg, _q=self.q: _q.put(("log", str(msg) + "\n"))

        t = threading.Thread(target=self._worker, args=(directory, density, opts, sheets, do_ann_shots),
                             daemon=True)
        t.start()

    # ---------- 后台任务 ----------
    def _worker(self, directory, density, opts, sheets, do_ann_shots):
        import pythoncom
        catia = None
        try:
            pythoncom.CoInitialize()
            catia = mc.connect_catia()
            files = sorted(f for f in os.listdir(directory)
                           if f.lower().endswith((".catpart", ".catproduct")))
            if not files:
                self.q.put(("done", "目录中没有 CATPart/CATProduct 文件:\n" + directory))
                return
            results = []
            all_texts = []
            ok_count = 0
            for fn in files:
                ok, text, data = mc.measure_one(catia, os.path.join(directory, fn), density, opts)
                results.append(data)
                if ok:
                    ok_count += 1
                all_texts.append(text)
                self.q.put(("log", text))

            # 三维标注逐个截图(可选)
            shots_map = {}
            if do_ann_shots:
                shots_dir = os.path.join(directory, "三维标注截图")
                part_shots_dir = os.path.join(directory, "零件截图")
                os.makedirs(shots_dir, exist_ok=True)
                os.makedirs(part_shots_dir, exist_ok=True)
                for dta in results:
                    if dta.get("kind") != "Part":
                        continue
                    part_path = os.path.join(directory, dta["file"])
                    if not os.path.exists(part_path):
                        continue
                    self.q.put(("log", "截图: " + dta["file"] + "\n"))
                    part_shot, all_shot, ann_results = mc.shots_for_part(
                        catia, part_path, shots_dir, dta.get("annotations") or [], part_shots_dir)
                    if part_shot:
                        dta["part_shot"] = part_shot
                    if all_shot:
                        dta["all_shot"] = all_shot
                    for ann_name, img in ann_results:
                        shots_map[(dta["file"], ann_name)] = img

            # 文本报告 / 中间数据 / Excel
            out_txt = os.path.join(directory, "测量结果.txt")
            with open(out_txt, "w", encoding="utf-8") as f:
                f.write("CATIA 测量结果  时间: {}\n目录: {}\n\n".format(
                    __import__("time").strftime("%Y-%m-%d %H:%M:%S"), directory))
                f.write("".join(all_texts))
            json_path = os.path.join(directory, "results.json")
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(results, f, ensure_ascii=False, indent=1)
            out_xlsx = os.path.join(directory, "测量结果.xlsx")
            mc.write_excel(out_xlsx, results, shots_map, sheets)
            self.q.put(("log", "Excel 已生成: " + out_xlsx + "\n"))
            self.q.put(("done", "完成: {}/{} 个文件处理成功\nExcel: {}".format(
                ok_count, len(files), out_xlsx)))
        except Exception as e:
            self.q.put(("done", "出错: {}".format(repr(e)[:300])))
        finally:
            try:
                if catia is not None:
                    catia.DisplayFileAlerts = True
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def run_gui():
    root = tk.Tk()
    App(root)
    root.mainloop()


def main():
    run_gui()


if __name__ == "__main__":
    main()
