
# -*- coding: utf-8 -*-
"""
Excel 合并助手（含简易GUI + 并行加速）
--------------------------------------------------
功能：
  1) 选择一个文件夹，将其中所有 .xlsx 文件读取并合并。
  2) 支持两种合并模式：
     - 合并为一个总表（所有文件/所有Sheet纵向堆叠到一个Sheet）
     - 按Sheet名分别合并（同名Sheet分别合并输出到多个Sheet）
  3) 可选项：递归子目录、添加来源文件/Sheet列、去空行、列并集/交集、去重、自动列宽等。
  4) 并行读取（线程池/进程池），在文件较多时能显著加速；UI线程保持响应。
  5) 超过 Excel 单表行上限（1,048,576）时自动分页写入 merged_1 / merged_2 ...。

依赖：pandas, openpyxl（写入也可选 xlsxwriter 更快）。

使用：
  直接运行本脚本，按界面提示操作。Windows 用户可用同目录下 run_windows.bat 自动创建虚拟环境并安装依赖。
"""
import os
import sys
import math
import json
import traceback
from pathlib import Path
from dataclasses import dataclass
from typing import List, Tuple, Dict, Optional, Any
from concurrent.futures import ThreadPoolExecutor, ProcessPoolExecutor, as_completed

# GUI
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 数据
import pandas as pd

EXCEL_MAX_ROWS = 1_048_576  # Excel 单表最大行数
TEMP_PREFIX = "~$"          # Excel 临时文件前缀，需跳过
CFG_FILE = "excel_merger_gui.config.json"

# -------- 工具函数 --------
def log_exception(e: BaseException) -> str:
    return "".join(traceback.format_exception(type(e), e, e.__traceback__))

def make_unique_columns(cols: List[Any]) -> List[str]:
    """为重复列名添加后缀 _1/_2，统一转字符串 & 去空白"""
    new_cols: List[str] = []
    seen: Dict[str, int] = {}
    for c in cols:
        name = str(c).strip()
        if name in seen:
            seen[name] += 1
            unique = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
            unique = name
        new_cols.append(unique)
    return new_cols

def safe_sheet_title(base: str, suffix: str = "", maxlen: int = 31) -> str:
    """Excel sheet 名称最长 31，必要时截断"""
    s = str(base)
    # Windows/Excel 不允许部分字符，这里做简单过滤
    for ch in ['\\', '/', '?', '*', '[', ']', ':']:
        s = s.replace(ch, '_')
    title = (s[: maxlen - len(suffix)] + suffix) if len(s) + len(suffix) > maxlen else s + suffix
    return title or "Sheet"

def chunk_dataframe(df: pd.DataFrame, limit: int = EXCEL_MAX_ROWS) -> List[pd.DataFrame]:
    n = len(df)
    if n <= limit:
        return [df]
    parts = []
    for i in range(0, n, limit):
        parts.append(df.iloc[i:i+limit].copy())
    return parts

def collect_xlsx_files(root: Path, recursive: bool = False) -> List[Path]:
    if recursive:
        files = [p for p in root.rglob("*.xlsx")]
    else:
        files = [p for p in root.glob("*.xlsx")]
    # 排除 Excel 临时文件 ~$.xlsx
    files = [p for p in files if not p.name.startswith(TEMP_PREFIX)]
    files.sort()
    return files

# -------- 并行 worker：读取单个文件所有Sheet --------
def read_excel_file_worker(
    file_path: str,
    drop_empty_rows: bool,
    add_src_file: bool,
    add_src_sheet: bool,
) -> Dict[str, Any]:
    """
    读取一个 .xlsx 文件的全部 sheet，返回：
    {
      "ok": bool,
      "file": "xxx.xlsx",
      "pairs": List[Tuple[str, pd.DataFrame]]  # [(sheet_name, df), ...]
      "err": Optional[str],
      "nrows": int
    }
    """
    fpath = Path(file_path)
    result: Dict[str, Any] = {"ok": True, "file": str(fpath), "pairs": [], "err": None, "nrows": 0}
    try:
        # 读取所有 Sheet：返回 dict[str, DataFrame]
        book = pd.read_excel(fpath, sheet_name=None, engine="openpyxl")
        for sname, df in book.items():
            if df is None:
                continue
            df = df.copy()
            # 标准化列名
            df.columns = make_unique_columns(list(df.columns))
            # 去掉全空行
            if drop_empty_rows:
                df.dropna(how="all", inplace=True)
            df.reset_index(drop=True, inplace=True)
            if add_src_file:
                df.insert(0, "source_file", fpath.name)
            if add_src_sheet:
                df.insert(1 if add_src_file else 0, "source_sheet", str(sname))
            result["pairs"].append((str(sname), df))
            result["nrows"] += len(df)
        return result
    except Exception as e:
        result["ok"] = False
        result["err"] = f"[读取失败] {fpath.name}: {e}\n{log_exception(e)}"
        return result

# -------- GUI 相关 --------
@dataclass
class Options:
    mode: str  # "single" | "by_sheet"
    recursive: bool
    add_src_file: bool
    add_src_sheet: bool
    drop_empty_rows: bool
    union_columns: bool  # True=并集, False=交集
    dedup_cols: List[str]
    parallel: bool
    parallel_kind: str    # "thread" | "process"
    max_workers: int
    auto_width: bool
    engine: str           # "openpyxl" | "xlsxwriter"

class MergeApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Excel 合并助手（.xlsx） - 带并行加速")
        self.geometry("860x640")
        self.minsize(860, 640)
        self._build_ui()
        self._load_cfg()
        self._working = False

    # ---------- UI ----------
    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}
        frm_top = ttk.Frame(self)
        frm_top.pack(fill="x", **pad)

        # 选择目录
        ttk.Label(frm_top, text="待合并文件夹：").grid(row=0, column=0, sticky="e")
        self.dir_var = tk.StringVar()
        self.dir_entry = ttk.Entry(frm_top, textvariable=self.dir_var, width=70)
        self.dir_entry.grid(row=0, column=1, sticky="we", padx=(0, 6))
        ttk.Button(frm_top, text="浏览...", command=self._browse_dir).grid(row=0, column=2, sticky="w")

        # 输出文件
        ttk.Label(frm_top, text="输出 .xlsx 文件：").grid(row=1, column=0, sticky="e")
        self.out_var = tk.StringVar()
        self.out_entry = ttk.Entry(frm_top, textvariable=self.out_var, width=70)
        self.out_entry.grid(row=1, column=1, sticky="we", padx=(0, 6))
        ttk.Button(frm_top, text="选择...", command=self._browse_save).grid(row=1, column=2, sticky="w")
        frm_top.columnconfigure(1, weight=1)

        # 合并模式
        frm_mode = ttk.LabelFrame(self, text="合并模式")
        frm_mode.pack(fill="x", **pad)
        self.mode_var = tk.StringVar(value="single")
        ttk.Radiobutton(frm_mode, text="合并为一个总表", variable=self.mode_var, value="single").grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Radiobutton(frm_mode, text="按Sheet名分别合并（同名Sheet各自合并）", variable=self.mode_var, value="by_sheet").grid(row=0, column=1, sticky="w", padx=8, pady=4)

        # 选项
        frm_opts = ttk.LabelFrame(self, text="选项")
        frm_opts.pack(fill="x", **pad)
        self.recursive_var = tk.BooleanVar(value=False)
        self.add_file_var = tk.BooleanVar(value=True)
        self.add_sheet_var = tk.BooleanVar(value=True)
        self.drop_empty_var = tk.BooleanVar(value=True)
        self.union_var = tk.BooleanVar(value=True)  # 默认并集
        self.auto_width_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm_opts, text="递归子文件夹", variable=self.recursive_var).grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(frm_opts, text="添加来源文件名列(source_file)", variable=self.add_file_var).grid(row=0, column=1, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(frm_opts, text="添加来源Sheet列(source_sheet)", variable=self.add_sheet_var).grid(row=0, column=2, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(frm_opts, text="去除空行", variable=self.drop_empty_var).grid(row=0, column=3, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(frm_opts, text="列并集（取消为交集）", variable=self.union_var).grid(row=1, column=0, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(frm_opts, text="自动列宽（前200行取样）", variable=self.auto_width_var).grid(row=1, column=1, sticky="w", padx=8, pady=4)

        ttk.Label(frm_opts, text="去重字段（逗号分隔，可留空）：").grid(row=2, column=0, sticky="e", padx=(8, 2))
        self.dedup_var = tk.StringVar(value="")
        ttk.Entry(frm_opts, textvariable=self.dedup_var, width=50).grid(row=2, column=1, columnspan=2, sticky="w", padx=2, pady=4)

        ttk.Label(frm_opts, text="写入引擎：").grid(row=2, column=3, sticky="e")
        self.engine_var = tk.StringVar(value="openpyxl")
        ttk.Combobox(frm_opts, textvariable=self.engine_var, values=["openpyxl", "xlsxwriter"], width=12, state="readonly").grid(row=2, column=4, sticky="w", padx=(2,8))

        # 并行
        frm_par = ttk.LabelFrame(self, text="并行加速")
        frm_par.pack(fill="x", **pad)
        self.parallel_var = tk.BooleanVar(value=True)
        self.kind_var = tk.StringVar(value="process")
        self.workers_var = tk.IntVar(value=max(2, (os.cpu_count() or 4) // 2))
        ttk.Checkbutton(frm_par, text="启用并行读取", variable=self.parallel_var).grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Label(frm_par, text="方式：").grid(row=0, column=1, sticky="e")
        ttk.Combobox(frm_par, textvariable=self.kind_var, values=["process", "thread"], width=10, state="readonly").grid(row=0, column=2, sticky="w", padx=(2,8))
        ttk.Label(frm_par, text="最大并发：").grid(row=0, column=3, sticky="e")
        ttk.Spinbox(frm_par, from_=1, to=max(64, (os.cpu_count() or 8)), textvariable=self.workers_var, width=6).grid(row=0, column=4, sticky="w", padx=(2,8))

        # 进度 & 按钮
        frm_run = ttk.Frame(self)
        frm_run.pack(fill="x", **pad)
        self.btn_start = ttk.Button(frm_run, text="开始合并", command=self._on_start)
        self.btn_start.pack(side="left")
        self.btn_stop = ttk.Button(frm_run, text="取消任务", command=self._on_stop, state="disabled")
        self.btn_stop.pack(side="left", padx=(8,0))
        self.prog = ttk.Progressbar(frm_run, mode="determinate")
        self.prog.pack(side="right", fill="x", expand=True)

        # 日志输出
        frm_log = ttk.LabelFrame(self, text="日志")
        frm_log.pack(fill="both", expand=True, **pad)
        self.txt = tk.Text(frm_log, height=18, wrap="word")
        sb = ttk.Scrollbar(frm_log, orient="vertical", command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        self.txt.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    # ---------- 配置 ----------
    def _load_cfg(self):
        try:
            cfg_path = Path(CFG_FILE)
            if cfg_path.exists():
                cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
                self.dir_var.set(cfg.get("dir", ""))
                self.out_var.set(cfg.get("out", ""))
        except Exception:
            pass

    def _save_cfg(self):
        try:
            cfg = {"dir": self.dir_var.get(), "out": self.out_var.get()}
            Path(CFG_FILE).write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    # ---------- 事件 ----------
    def _browse_dir(self):
        d = filedialog.askdirectory(title="选择包含 .xlsx 的文件夹")
        if d:
            self.dir_var.set(d)

    def _browse_save(self):
        f = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if f:
            self.out_var.set(f)

    def _on_stop(self):
        self._working = False
        self.log("收到取消请求，正在尝试终止...（当前批次完成后停止）")
        self.btn_stop.configure(state="disabled")

    def _on_start(self):
        if self._working:
            return
        in_dir = self.dir_var.get().strip()
        out_file = self.out_var.get().strip()
        if not in_dir:
            messagebox.showwarning("提示", "请选择待合并文件夹")
            return
        if not out_file:
            messagebox.showwarning("提示", "请选择输出 .xlsx 文件路径")
            return
        root = Path(in_dir)
        if not root.exists() or not root.is_dir():
            messagebox.showerror("错误", "输入目录不存在或不是文件夹")
            return
        # 组装选项
        opts = Options(
            mode=self.mode_var.get(),
            recursive=self.recursive_var.get(),
            add_src_file=self.add_file_var.get(),
            add_src_sheet=self.add_sheet_var.get(),
            drop_empty_rows=self.drop_empty_var.get(),
            union_columns=self.union_var.get(),
            dedup_cols=[c.strip() for c in self.dedup_var.get().split(",") if c.strip()],
            parallel=self.parallel_var.get(),
            parallel_kind=self.kind_var.get(),
            max_workers=max(1, int(self.workers_var.get() or 1)),
            auto_width=self.auto_width_var.get(),
            engine=self.engine_var.get()
        )
        self._save_cfg()
        # 启动任务
        self._working = True
        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.txt.delete("1.0", "end")
        self.prog.configure(value=0, maximum=100)
        self.after(100, lambda: self._run_merge(root, Path(out_file), opts))

    # ---------- 日志 & 进度 ----------
    def log(self, msg: str):
        self.txt.insert("end", msg + "\n")
        self.txt.see("end")
        self.update_idletasks()

    def set_progress(self, value: float, maximum: Optional[float] = None):
        if maximum is not None:
            self.prog.configure(maximum=maximum)
        self.prog.configure(value=value)
        self.update_idletasks()

    # ---------- 主流程 ----------
    def _run_merge(self, root: Path, out_file: Path, opts: Options):
        try:
            files = collect_xlsx_files(root, recursive=opts.recursive)
            # 排除输出文件（若位于同一目录）
            files = [f for f in files if f.resolve() != out_file.resolve()]
            if not files:
                self.log("未在该目录下找到 .xlsx 文件。")
                messagebox.showinfo("提示", "未找到 .xlsx 文件")
                return self._finish()
            self.log(f"发现 {len(files)} 个 .xlsx 文件，即将开始解析...")

            # 读取阶段进度
            self.set_progress(0, maximum=len(files) + 1)

            all_pairs: List[Tuple[str, pd.DataFrame]] = []
            read_ok = 0
            # 并行 or 串行读取
            if opts.parallel:
                max_workers = opts.max_workers
                if opts.parallel_kind == "process":
                    ex_cls = ProcessPoolExecutor
                else:
                    ex_cls = ThreadPoolExecutor
                self.log(f"并行模式：{opts.parallel_kind}，并发数：{max_workers}")
                with ex_cls(max_workers=max_workers) as ex:
                    fut_map = {}
                    for f in files:
                        if not self._working:
                            break
                        fut = ex.submit(
                            read_excel_file_worker,
                            str(f), opts.drop_empty_rows, opts.add_src_file, opts.add_src_sheet
                        )
                        fut_map[fut] = f
                    for fut in as_completed(fut_map):
                        if not self._working:
                            break
                        f = fut_map[fut]
                        try:
                            res = fut.result()
                        except Exception as e:
                            self.log(f"[异常] {f.name}: {e}")
                            continue
                        if res["ok"]:
                            read_ok += 1
                            all_pairs.extend(res["pairs"])
                            self.log(f"[读取完成] {Path(res['file']).name}（{len(res['pairs'])} 个Sheet，共 {res['nrows']} 行）")
                        else:
                            self.log(res["err"] or f"[读取失败] {f.name}")
                        self.set_progress(read_ok, maximum=len(files) + 1)
            else:
                for i, f in enumerate(files, 1):
                    if not self._working:
                        break
                    res = read_excel_file_worker(
                        str(f), opts.drop_empty_rows, opts.add_src_file, opts.add_src_sheet
                    )
                    if res["ok"]:
                        read_ok += 1
                        all_pairs.extend(res["pairs"])
                        self.log(f"[读取完成] {Path(res['file']).name}（{len(res['pairs'])} 个Sheet，共 {res['nrows']} 行）")
                    else:
                        self.log(res["err"] or f"[读取失败] {f.name}")
                    self.set_progress(i, maximum=len(files) + 1)

            if not self._working:
                self.log("任务已取消。")
                return self._finish()

            if not all_pairs:
                self.log("未读取到任何数据，终止。")
                messagebox.showwarning("提示", "未读取到任何数据")
                return self._finish()

            # 合并阶段
            self.log(f"开始合并（模式：{'总表' if opts.mode=='single' else '按Sheet'}，列集合：{'并集' if opts.union_columns else '交集'}）...")
            out_file.parent.mkdir(parents=True, exist_ok=True)

            with pd.ExcelWriter(out_file, engine=opts.engine) as writer:
                written_sheets = 0
                if opts.mode == "single":
                    dfs = [df for (_s, df) in all_pairs]
                    if not dfs:
                        self.log("没有可合并的数据帧。")
                    else:
                        if opts.union_columns:
                            merged = pd.concat(dfs, ignore_index=True, sort=False)
                        else:
                            # 交集列
                            commons = None
                            for d in dfs:
                                s = set(d.columns)
                                commons = s if commons is None else commons & s
                            commons = list(commons) if commons else []
                            merged = pd.concat([d[commons].copy() for d in dfs], ignore_index=True)

                        # 去重
                        if opts.dedup_cols:
                            try:
                                merged.drop_duplicates(subset=opts.dedup_cols, inplace=True, ignore_index=True)
                            except Exception as e:
                                self.log(f"[去重警告] {e}")

                        parts = chunk_dataframe(merged, EXCEL_MAX_ROWS)
                        for idx, part in enumerate(parts, 1):
                            sheet_name = "merged" if len(parts) == 1 else f"merged_{idx}"
                            part.to_excel(writer, sheet_name=sheet_name, index=False)
                            if opts.auto_width and hasattr(writer, "book"):
                                try:
                                    self._adjust_width(writer, sheet_name, part)
                                except Exception:
                                    pass
                            written_sheets += 1
                            self.log(f"[写入] {sheet_name}  行数：{len(part)}")
                else:
                    # 按 sheet name 分组合并
                    groups: Dict[str, List[pd.DataFrame]] = {}
                    for s, df in all_pairs:
                        groups.setdefault(s, []).append(df)
                    for s, lst in groups.items():
                        if opts.union_columns:
                            merged = pd.concat(lst, ignore_index=True, sort=False)
                        else:
                            commons = None
                            for d in lst:
                                sset = set(d.columns)
                                commons = sset if commons is None else commons & sset
                            commons = list(commons) if commons else []
                            merged = pd.concat([d[commons].copy() for d in lst], ignore_index=True)

                        if opts.dedup_cols:
                            try:
                                merged.drop_duplicates(subset=opts.dedup_cols, inplace=True, ignore_index=True)
                            except Exception as e:
                                self.log(f"[去重警告-{s}] {e}")

                        parts = chunk_dataframe(merged, EXCEL_MAX_ROWS)
                        for idx, part in enumerate(parts, 1):
                            suffix = "" if len(parts) == 1 else f"_{idx}"
                            sheet_name = safe_sheet_title(s, suffix=suffix)
                            part.to_excel(writer, sheet_name=sheet_name, index=False)
                            if opts.auto_width and hasattr(writer, "book"):
                                try:
                                    self._adjust_width(writer, sheet_name, part)
                                except Exception:
                                    pass
                            written_sheets += 1
                            self.log(f"[写入] {sheet_name}  行数：{len(part)}")

            self.set_progress(len(files) + 1, maximum=len(files) + 1)
            self.log(f"✅ 合并完成，输出：{out_file}")
            try:
                # Windows 下合并后可一键打开所在目录
                if sys.platform.startswith("win") and out_file.exists():
                    os.startfile(out_file.parent)  # type: ignore
            except Exception:
                pass
            messagebox.showinfo("完成", f"合并完成！\n输出文件：\n{out_file}")
        except Exception as e:
            self.log("发生异常：\n" + log_exception(e))
            messagebox.showerror("错误", str(e))
        finally:
            self._finish()

    def _finish(self):
        self._working = False
        self.btn_start.configure(state="normal")
        self.btn_stop.configure(state="disabled")

    def _adjust_width(self, writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame, sample_rows: int = 200):
        """基于前 sample_rows 行估算列宽（openpyxl/xlsxwriter）"""
        try:
            if writer.engine == "openpyxl":
                from openpyxl.utils import get_column_letter
                ws = writer.book[sheet_name] if hasattr(writer.book, "__getitem__") else writer.sheets[sheet_name]
                for i, col in enumerate(df.columns, start=1):
                    series = df[col].astype(str).head(sample_rows)
                    max_len = max([len(str(col))] + [len(v) for v in series.tolist()]) if len(series) else len(str(col))
                    max_len = min(60, max_len + 2)
                    ws.column_dimensions[get_column_letter(i)].width = max_len
                # 冻结首行
                ws.freeze_panes = "A2"
            elif writer.engine == "xlsxwriter":
                ws = writer.sheets[sheet_name]
                for i, col in enumerate(df.columns):
                    series = df[col].astype(str).head(sample_rows)
                    max_len = max([len(str(col))] + [len(v) for v in series.tolist()]) if len(series) else len(str(col))
                    max_len = min(60, max_len + 2)
                    ws.set_column(i, i, max_len)
                ws.freeze_panes(1, 0)
        except Exception:
            # 调整失败也不影响主流程
            pass

def main():
    app = MergeApp()
    app.mainloop()

if __name__ == "__main__":
    # Windows 进程池所需
    try:
        from multiprocessing import freeze_support
        freeze_support()
    except Exception:
        pass
    main()
