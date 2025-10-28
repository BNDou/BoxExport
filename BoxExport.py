#!/usr/bin/env python
# coding=utf-8
'''
FilePath     : /BoxExport/BoxExport.py
Description  :  
Author       : BNDou
Date         : 2025-10-28 21:32:19
LastEditTime : 2025-10-28 23:20:22
'''

import os
import sys
import time
import math
import threading
import webbrowser
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk
from typing import List, Tuple, Optional, Callable, Dict

try:
    import xlrd
except ImportError as e:
    print("缺少依赖 xlrd，请先安装：pip install xlrd==1.2.0", file=sys.stderr)
    raise

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, Border, Side
except ImportError as e:
    print("缺少依赖 openpyxl，请先安装：pip install openpyxl", file=sys.stderr)
    raise

# PyInstaller/auto-py-to-exe 打包时，openpyxl 某些内部 writer 模块可能不会被自动收集。
# 下面的显式导入用于作为打包钩子，运行时若不存在也不会报错。
try:
    import openpyxl.cell._writer  # type: ignore
    import openpyxl.workbook._writer  # type: ignore
    import openpyxl.worksheet._writer  # type: ignore
except Exception:
    pass


def ensure_output_dir(script_dir: str) -> str:
    output_dir = os.path.join(script_dir, "output")
    if not os.path.isdir(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    return output_dir


def pick_data_directory(initial_dir: Optional[str] = None) -> Optional[str]:
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askdirectory(title="选择数据源文件夹 data", initialdir=initial_dir or os.getcwd())
    root.update()
    if not path:
        return None
    return path


def list_xls_files(folder: str) -> List[str]:
    files = []
    for name in os.listdir(folder):
        if name.lower().endswith(".xls") and not name.startswith("~$"):
            files.append(os.path.join(folder, name))
    # 按文件名中的数字顺序排序（例如 000013.XLS -> 13）
    def key_func(p: str) -> Tuple[int, str]:
        base = os.path.basename(p)
        stem = os.path.splitext(base)[0]
        try:
            return (int(stem.lstrip('0') or '0'), base)
        except ValueError:
            return (10**9, base)

    files.sort(key=key_func)
    return files


def xcell_value(sheet: "xlrd.sheet.Sheet", row: int, col: int):
    cell = sheet.cell(row, col)
    if cell.ctype == xlrd.XL_CELL_DATE:
        try:
            dt = xlrd.xldate_as_datetime(cell.value, sheet.book.datemode)
            return dt
        except Exception:
            return cell.value
    return cell.value


def read_records_from_xls(xls_path: str) -> List[dict]:
    """
    源文件格式（按列）：
    A 归档号（不导出）
    C 病案号
    D 出院科室
    E 患者姓名
    F 出院时间
    G 入院时间
    H 成像张数
    I 备注（暂不导出）

    假设第2行是表头，从第3行开始为数据；遇到空的“病案号”即停止。
    """
    book = xlrd.open_workbook(xls_path)
    sheet = book.sheet_by_index(0)

    records: List[dict] = []

    # 从第3行开始（0-based -> 行索引2）
    start_row = 2
    for r in range(start_row, sheet.nrows):
        case_no = xcell_value(sheet, r, 2)  # C列（index=2）
        if case_no in (None, ""):
            # 若病案号空，认为到达数据尾部
            continue

        department = xcell_value(sheet, r, 3)  # D
        patient_name = xcell_value(sheet, r, 4)  # E
        discharge_time = xcell_value(sheet, r, 5)  # F
        admission_time = xcell_value(sheet, r, 6)  # G
        image_count = xcell_value(sheet, r, 7)  # H

        records.append({
            "case_no": case_no,
            "department": department,
            "patient_name": patient_name,
            "discharge_time": discharge_time,
            "admission_time": admission_time,
            "image_count": image_count,
        })

    return records


def extract_box_number_from_filename(path: str) -> Optional[int]:
    base = os.path.basename(path)
    stem = os.path.splitext(base)[0]
    try:
        return int(stem.lstrip('0') or '0')
    except ValueError:
        return None


def build_and_export_workbook(records: List[dict], box_number: Optional[int], sequence_start: int, out_path: str) -> int:
    wb = Workbook()
    ws = wb.active

    center = Alignment(horizontal="center", vertical="center")
    bold_center = Alignment(horizontal="center", vertical="center")
    bold_font = Font(bold=True)
    thin_side = Side(style="thin", color="000000")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # A1-H1 合并，标题
    ws.merge_cells("A1:H1")
    ws["A1"] = "卷内目录"
    ws["A1"].alignment = center
    ws["A1"].font = Font(name="宋体", size=20, bold=True)
    ws.row_dimensions[1].height = 25

    # E2, F2, G2, H2 固定内容
    ws["E2"] = "全宗号"
    ws["F2"] = "120"
    ws["G2"] = "类目"
    ws["H2"] = "BL"
    for addr in ("E2", "F2", "G2", "H2"):
        ws[addr].alignment = center
        ws[addr].font = Font(name="宋体", size=16)
    ws.row_dimensions[2].height = 25

    # G3, H3 箱号
    ws["G3"] = "箱号"
    ws["G3"].alignment = center
    ws["G3"].font = Font(name="宋体", size=16)
    ws["H3"] = box_number if box_number is not None else ""
    ws["H3"].alignment = center
    ws["H3"].font = Font(name="宋体", size=16)
    ws.row_dimensions[3].height = 25

    # 第4行表头
    headers = [
        ("A4", "顺序号"),
        ("B4", "病案号"),
        ("C4", "出院科室"),
        ("D4", "患者姓名"),
        ("E4", "出院时间"),
        ("F4", "入院时间"),
        ("G4", "成像张数"),
        ("H4", "备注"),
    ]
    for addr, text in headers:
        ws[addr] = text
        ws[addr].alignment = bold_center
        ws[addr].font = bold_font

    # 列宽（按要求）
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 11
    ws.column_dimensions['C'].width = 13
    ws.column_dimensions['D'].width = 8
    ws.column_dimensions['E'].width = 11
    ws.column_dimensions['F'].width = 11
    ws.column_dimensions['G'].width = 8
    ws.column_dimensions['H'].width = 15

    # 数据起始行
    current_row = 5
    seq = sequence_start

    for rec in records:
        # A 顺序号（居中）
        ws.cell(row=current_row, column=1, value=seq).alignment = center

        # B-H 数据
        ws.cell(row=current_row, column=2, value=rec.get("case_no"))
        ws.cell(row=current_row, column=3, value=rec.get("department"))
        ws.cell(row=current_row, column=4, value=rec.get("patient_name"))

        # E 出院时间, F 入院时间（尽量保持日期展示）
        e_val = rec.get("discharge_time")
        f_val = rec.get("admission_time")
        e_cell = ws.cell(row=current_row, column=5, value=e_val)
        f_cell = ws.cell(row=current_row, column=6, value=f_val)
        # 若为 datetime，设置日期格式
        try:
            from datetime import datetime
            if isinstance(e_val, datetime):
                e_cell.number_format = "yyyy-mm-dd"
            if isinstance(f_val, datetime):
                f_cell.number_format = "yyyy-mm-dd"
        except Exception:
            pass

        # G 成像张数（居中）
        g_cell = ws.cell(row=current_row, column=7, value=rec.get("image_count"))
        g_cell.alignment = center

        # H 备注 先留空
        ws.cell(row=current_row, column=8, value="")

        current_row += 1
        seq += 1

    # 添加表格边框：A1 到 H{最后一行}
    last_row = ws.max_row
    for r in range(1, last_row + 1):
        for c in range(1, 8 + 1):
            ws.cell(row=r, column=c).border = thin_border

    wb.save(out_path)
    # 返回下一个可用顺序号
    return seq


def process_all(
    data_dir: str,
    sequence_start: int,
    output_dir: str,
    on_progress: Optional[Callable[[int, int, str, Dict[str, float]], None]] = None,
    on_log: Optional[Callable[[str], None]] = None,
    should_stop: Optional[Callable[[], bool]] = None,
) -> None:
    # 确保用户自定义的导出目录存在
    os.makedirs(output_dir, exist_ok=True)

    xls_files = list_xls_files(data_dir)
    if not xls_files:
        raise RuntimeError("所选目录未找到 .XLS 文件。")

    current_seq = sequence_start
    total = len(xls_files)
    start_time = time.time()

    def ftime(sec: float) -> str:
        sec = int(sec)
        h = sec // 3600
        m = (sec % 3600) // 60
        s = sec % 60
        if h > 0:
            return f"{h:02d}:{m:02d}:{s:02d}"
        else:
            return f"{m:02d}:{s:02d}"

    for idx, xls in enumerate(xls_files, start=1):
        file_start = time.time()
        name = os.path.basename(xls)
        if on_log:
            on_log(f"读取: {name}")
        records = read_records_from_xls(xls)

        box_no = extract_box_number_from_filename(xls)
        out_name = f"{os.path.splitext(name)[0]}.xlsx"
        out_path = os.path.join(output_dir, out_name)

        if on_log:
            on_log(f"导出: {out_name}")
        next_seq = build_and_export_workbook(records, box_no, current_seq, out_path)
        current_seq = next_seq

        file_elapsed = time.time() - file_start
        elapsed = time.time() - start_time
        avg_per_file = elapsed / idx
        remain_files = total - idx
        eta_seconds = max(0.0, avg_per_file * remain_files)
        if on_progress:
            on_progress(
                idx,
                total,
                name,
                {
                    "file_elapsed": file_elapsed,
                    "elapsed": elapsed,
                    "avg_per_file": avg_per_file,
                    "eta_seconds": eta_seconds,
                },
            )

        # 若用户请求停止，则在完成当前文件后终止后续处理
        if should_stop and should_stop():
            if on_log:
                on_log("收到停止指令，已在当前文件完成后终止后续处理。")
            break

    if on_log:
        on_log(f"完成，输出目录: {output_dir}")


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("病案卷内目录批量生成")
        self.root.geometry("820x520+0+0")

        self.data_dir_var = tk.StringVar()
        self.seq_var = tk.IntVar(value=1)
        self.output_dir_var = tk.StringVar(value=os.path.join(os.path.dirname(os.path.abspath(__file__)), "output"))
        self.status_var = tk.StringVar(value="就绪")
        self.percent_var = tk.StringVar(value="0.00%")
        self.eta_var = tk.StringVar(value="ETA 00:00")
        self.elapsed_var = tk.StringVar(value="已用 00:00")
        self.avg_var = tk.StringVar(value="平均 00:00")

        self._build_ui()

        self._total = 0
        self._idx = 0
        self._start_ts = 0.0
        self._running = False

    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}

        frm_top = ttk.Frame(self.root)
        frm_top.pack(fill=tk.X, **pad)

        ttk.Label(frm_top, text="数据目录:").pack(side=tk.LEFT)
        ent_dir = ttk.Entry(frm_top, textvariable=self.data_dir_var, width=70)
        ent_dir.pack(side=tk.LEFT, padx=6)
        ttk.Button(frm_top, text="浏览…", command=self._browse).pack(side=tk.LEFT)

        frm_seq = ttk.Frame(self.root)
        frm_seq.pack(fill=tk.X, **pad)
        ttk.Label(frm_seq, text="起始顺序号:").pack(side=tk.LEFT)
        spn = ttk.Spinbox(frm_seq, from_=1, to=999999, textvariable=self.seq_var, width=8)
        spn.pack(side=tk.LEFT, padx=6)
        self.btn_start = ttk.Button(frm_seq, text="开始", command=self._start)
        self.btn_start.pack(side=tk.LEFT)
        self.btn_stop = ttk.Button(frm_seq, text="停止", command=self._stop, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=6)

        frm_out = ttk.Frame(self.root)
        frm_out.pack(fill=tk.X, **pad)
        ttk.Label(frm_out, text="导出目录:").pack(side=tk.LEFT)
        ent_out = ttk.Entry(frm_out, textvariable=self.output_dir_var, width=70)
        ent_out.pack(side=tk.LEFT, padx=6)
        ttk.Button(frm_out, text="浏览…", command=self._browse_out).pack(side=tk.LEFT)
        ttk.Button(frm_out, text="打开导出目录", command=self._open_out_dir).pack(side=tk.LEFT, padx=6)

        frm_prog = ttk.Frame(self.root)
        frm_prog.pack(fill=tk.X, **pad)
        self.prog = ttk.Progressbar(frm_prog, orient=tk.HORIZONTAL, length=600, mode='determinate', maximum=100)
        self.prog.pack(side=tk.LEFT, expand=True, fill=tk.X)
        ttk.Label(frm_prog, textvariable=self.percent_var, width=8).pack(side=tk.LEFT, padx=6)

        frm_stats = ttk.Frame(self.root)
        frm_stats.pack(fill=tk.X, **pad)
        ttk.Label(frm_stats, textvariable=self.status_var).pack(side=tk.LEFT)
        ttk.Label(frm_stats, textvariable=self.elapsed_var).pack(side=tk.LEFT, padx=12)
        ttk.Label(frm_stats, textvariable=self.avg_var).pack(side=tk.LEFT, padx=12)
        ttk.Label(frm_stats, textvariable=self.eta_var).pack(side=tk.LEFT, padx=12)

        frm_log = ttk.LabelFrame(self.root, text="日志")
        frm_log.pack(fill=tk.BOTH, expand=True, **pad)
        self.txt = tk.Text(frm_log, height=16)
        self.txt.pack(fill=tk.BOTH, expand=True)

        # 右下角版本与作者信息，可点击打开链接
        frm_bottom = ttk.Frame(self.root)
        frm_bottom.pack(fill=tk.X, side=tk.BOTTOM)
        link = tk.Label(
            frm_bottom,
            text="v2025.10.28 by BNDou",
            fg="blue",
            font=("SimHei", 10, "bold"),
            cursor="hand2",
            anchor=tk.E,
        )
        link.pack(side=tk.RIGHT, padx=10, pady=6)
        link.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/BNDou/BoxExport"))

    def _browse(self):
        d = filedialog.askdirectory(title="选择数据源文件夹 data", initialdir=self.data_dir_var.get() or os.getcwd())
        if d:
            self.data_dir_var.set(d)

    def _browse_out(self):
        d = filedialog.askdirectory(title="选择导出目录", initialdir=self.output_dir_var.get() or os.getcwd())
        if d:
            self.output_dir_var.set(d)

    def _open_out_dir(self):
        p = self.output_dir_var.get().strip()
        if not p:
            messagebox.showwarning("提示", "请先选择导出目录。")
            return
        if not os.path.isdir(p):
            messagebox.showwarning("提示", "导出目录不存在，请先生成或选择有效目录。")
            return
        try:
            if sys.platform.startswith('win'):
                os.startfile(p)  # type: ignore
            elif sys.platform == 'darwin':
                subprocess.call(['open', p])
            else:
                subprocess.call(['xdg-open', p])
        except Exception as e:
            messagebox.showerror("错误", f"无法打开目录：{e}")

    def _start(self):
        if self._running:
            return
        data_dir = self.data_dir_var.get().strip()
        if not data_dir or not os.path.isdir(data_dir):
            messagebox.showwarning("提示", "请先选择有效的数据目录。")
            return
        output_dir = self.output_dir_var.get().strip()
        if not output_dir:
            messagebox.showwarning("提示", "请先选择导出目录。")
            return
        seq = int(self.seq_var.get() or 1)
        self._running = True
        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self.prog['value'] = 0
        self.percent_var.set("0.00%")
        self.txt.delete(1.0, tk.END)
        self.status_var.set("开始处理…")
        self._start_ts = time.time()
        self._stop_requested = False

        def run():
            try:
                process_all(
                    data_dir,
                    seq,
                    output_dir,
                    on_progress=self._on_progress,
                    on_log=self._on_log,
                    should_stop=lambda: self._stop_requested,
                )
                self.root.after(0, lambda: messagebox.showinfo("完成", "处理完成。请在导出目录查看导出文件。"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
            finally:
                def finish():
                    self._running = False
                    self.btn_start.config(state=tk.NORMAL)
                    self.btn_stop.config(state=tk.DISABLED)
                    self.status_var.set("就绪")
                    self._stop_requested = False
                self.root.after(0, finish)

        threading.Thread(target=run, daemon=True).start()

    def _on_log(self, msg: str):
        def do():
            self.txt.insert(tk.END, msg + "\n")
            self.txt.see(tk.END)
        self.root.after(0, do)

    def _on_progress(self, idx: int, total: int, name: str, stats: Dict[str, float]):
        percent = idx * 100.0 / total if total else 0.0
        def ftime(sec: float) -> str:
            sec = int(sec)
            h = sec // 3600
            m = (sec % 3600) // 60
            s = sec % 60
            if h > 0:
                return f"{h:02d}:{m:02d}:{s:02d}"
            else:
                return f"{m:02d}:{s:02d}"
        def do():
            self.prog['value'] = percent
            self.percent_var.set(f"{percent:0.2f}%")
            self.status_var.set(f"{idx}/{total}  当前: {name}")
            self.elapsed_var.set(f"已用 {ftime(stats.get('elapsed', 0.0))}")
            self.avg_var.set(f"平均 {ftime(stats.get('avg_per_file', 0.0))}")
            self.eta_var.set(f"ETA {ftime(stats.get('eta_seconds', 0.0))}")
        self.root.after(0, do)

    def _stop(self):
        if not self._running:
            return
        self._stop_requested = True
        self.status_var.set("已请求停止：将完成当前文件后终止…")


def main_gui():
    root = tk.Tk()
    App(root)
    root.mainloop()


def _wait_exit():
    try:
        print("输入 0 然后回车以退出程序...")
        while True:
            user_in = input().strip()
            if user_in == "0":
                break
    except Exception:
        pass


if __name__ == "__main__":
    # 默认启动 GUI
    main_gui()


