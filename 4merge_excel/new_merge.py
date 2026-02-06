# -*- coding: utf-8 -*-
"""
Excel 合并助手（.xlsx） - 简易GUI + 并行加速（改进版）
--------------------------------------------------
本版修复 / 增强：
1) 输出路径支持“选择文件夹”：若你选择的是文件夹，将自动生成 merged.xlsx（可自行改名）。
2) 修复 PermissionError 常见原因：
   - 输出路径指向文件夹（没有文件名） -> 自动补全为文件夹/merged.xlsx
   - 输出文件被 Excel 占用（打开未关闭） -> 给出明确提示
3) 其他增强：
   - 去重字段支持：逗号/中文逗号/分号/换行分隔
   - Sheet 名称去重（避免 31 字截断后重名）
   - 写入采用“临时文件 -> 原子替换”，尽量避免中途失败留下半成品
   - xlsxwriter 不可用时自动回退到 openpyxl 并提示
依赖：pandas, openpyxl（可选 xlsxwriter）
"""
from __future__ import annotations

import os
import sys
import json
import re
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
    """Excel sheet 名称最长 31，过滤非法字符并按需截断"""
    s = str(base).strip()
    for ch in ['\\', '/', '?', '*', '[', ']', ':']:
        s = s.replace(ch, '_')
    if not s:
        s = "Sheet"
    # 截断
    if len(s) + len(suffix) > maxlen:
        s = s[: max(1, maxlen - len(suffix))]
    return s + suffix


def ensure_unique_sheet_name(name: str, used: set[str]) -> str:
    """避免 sheet 重名（特别是 31 字截断导致的重名）"""
    if name not in used:
        used.add(name)
        return name
    base = name
    # 尝试 _2/_3...
    for i in range(2, 10_000):
        suffix = f"_{i}"
        candidate = safe_sheet_title(base, suffix=suffix)
        if candidate not in used:
            used.add(candidate)
            return candidate
    # 极端情况
    raise RuntimeError(f"Sheet 名称去重失败：{name}")


def chunk_dataframe(df: pd.DataFrame, limit: int = EXCEL_MAX_ROWS) -> List[pd.DataFrame]:
    n = len(df)
    if n <= limit:
        return [df]
    return [df.iloc[i:i + limit].copy() for i in range(0, n, limit)]


def collect_xlsx_files(root: Path, recursive: bool = False) -> List[Path]:
    files = list(root.rglob("*.xlsx")) if recursive else list(root.glob("*.xlsx"))
    files = [p for p in files if not p.name.startswith(TEMP_PREFIX)]
    files.sort()
    return files


def parse_fields(text: str) -> List[str]:
    """解析“去重字段”输入：支持逗号/中文逗号/分号/换行"""
    if not text:
        return []
    parts = re.split(r"[,\uFF0C;\uFF1B\r\n]+", text)
    return [p.strip() for p in parts if p.strip()]


def resolve_output_path(raw: str, default_filename: str = "merged.xlsx") -> Path:
    """
    允许用户输入：
    - 完整文件路径：E:\\out\\abc.xlsx
    - 不带扩展名：E:\\out\\abc  -> 自动补 .xlsx
    - 文件夹路径：E:\\out\\folder -> 自动变为 folder\\merged.xlsx
    """
    s = (raw or "").strip().strip('"').strip("'")
    if not s:
        raise ValueError("输出路径不能为空")
    p = Path(s)

    # 情况1：已存在且是目录 -> 目录/默认名
    if p.exists() and p.is_dir():
        return (p / default_filename)

    # 情况2：明显像目录（末尾是分隔符）
    if s.endswith(("\\", "/")):
        return (p / default_filename)

    # 情况3：没有后缀 -> 认为是“文件名但没写 .xlsx”，补上
    if p.suffix == "":
        p = p.with_suffix(".xlsx")
        return p

    # 情况4：有后缀但不是 .xlsx -> 也强制改为 .xlsx（避免写入失败）
    if p.suffix.lower() != ".xlsx":
        p = p.with_suffix(".xlsx")
    return p


def pick_writer_engine(engine: str) -> str:
    """xlsxwriter 不可用时回退 openpyxl"""
    eng = (engine or "openpyxl").lower().strip()
    if eng == "xlsxwriter":
        try:
            import xlsxwriter  # noqa: F401
            return "xlsxwriter"
        except Exception:
            return "openpyxl"
    return "openpyxl"


def atomic_replace(src: Path, dst: Path) -> None:
    """
    尽量原子替换：先写到临时文件，再替换目标文件。
    Windows 下若目标被占用，会在 replace/rename 抛 PermissionError。
    """
    try:
        # Python 3.8+：Path.replace 是 os.replace（原子）
        src.replace(dst)
    except Exception:
        # 兜底：先删后改名
        if dst.exists():
            dst.unlink()
        src.replace(dst)


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
        book = pd.read_excel(fpath, sheet_name=None, engine="openpyxl")
        for sname, df in book.items():
            if df is None:
                continue
            df = df.copy()
            df.columns = make_unique_columns(list(df.columns))
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
        self.geometry("920x680")
        self.minsize(920, 680)
        self._working = False
        self._build_ui()
        self._load_cfg()

    # ---------- UI ----------
    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}
        frm_top = ttk.Frame(self)
        frm_top.pack(fill="x", **pad)

        # 选择目录
        ttk.Label(frm_top, text="待合并文件夹：").grid(row=0, column=0, sticky="e")
        self.dir_var = tk.StringVar()
        self.dir_entry = ttk.Entry(frm_top, textvariable=self.dir_var, width=72)
        self.dir_entry.grid(row=0, column=1, sticky="we", padx=(0, 6))
        ttk.Button(frm_top, text="浏览...", command=self._browse_dir).grid(row=0, column=2, sticky="w")

        # 输出路径（支持文件/文件夹）
        ttk.Label(frm_top, text="输出路径（文件或文件夹）：").grid(row=1, column=0, sticky="e")
        self.out_var = tk.StringVar()
        self.out_entry = ttk.Entry(frm_top, textvariable=self.out_var, width=72)
        self.out_entry.grid(row=1, column=1, sticky="we", padx=(0, 6))
        ttk.Button(frm_top, text="选文件...", command=self._browse_save).grid(row=1, column=2, sticky="w")
        ttk.Button(frm_top, text="选文件夹...", command=self._browse_out_dir).grid(row=1, column=3, sticky="w", padx=(6, 0))
        ttk.Button(frm_top, text="打开目录", command=self._open_out_dir).grid(row=1, column=4, sticky="w", padx=(6, 0))

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
        self.union_var = tk.BooleanVar(value=True)
        self.auto_width_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm_opts, text="递归子文件夹", variable=self.recursive_var).grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(frm_opts, text="添加来源文件名列(source_file)", variable=self.add_file_var).grid(row=0, column=1, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(frm_opts, text="添加来源Sheet列(source_sheet)", variable=self.add_sheet_var).grid(row=0, column=2, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(frm_opts, text="去除空行", variable=self.drop_empty_var).grid(row=0, column=3, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(frm_opts, text="列并集（取消为交集）", variable=self.union_var).grid(row=1, column=0, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(frm_opts, text="自动列宽（前200行取样）", variable=self.auto_width_var).grid(row=1, column=1, sticky="w", padx=8, pady=4)

        ttk.Label(frm_opts, text="去重字段（逗号/分号/换行分隔，可留空）：").grid(row=2, column=0, sticky="e", padx=(8, 2))
        self.dedup_var = tk.StringVar(value="")
        ttk.Entry(frm_opts, textvariable=self.dedup_var, width=54).grid(row=2, column=1, columnspan=2, sticky="w", padx=2, pady=4)

        ttk.Label(frm_opts, text="写入引擎：").grid(row=2, column=3, sticky="e")
        self.engine_var = tk.StringVar(value="openpyxl")
        ttk.Combobox(frm_opts, textvariable=self.engine_var, values=["openpyxl", "xlsxwriter"], width=12, state="readonly").grid(row=2, column=4, sticky="w", padx=(2, 8))

        # 并行
        frm_par = ttk.LabelFrame(self, text="并行加速")
        frm_par.pack(fill="x", **pad)
        self.parallel_var = tk.BooleanVar(value=True)
        # 读 xlsx 通常 IO+解压，线程常更稳；进程会复制/序列化 DataFrame，可能更吃内存
        self.kind_var = tk.StringVar(value="thread")
        self.workers_var = tk.IntVar(value=max(2, (os.cpu_count() or 4) // 2))
        ttk.Checkbutton(frm_par, text="启用并行读取", variable=self.parallel_var).grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Label(frm_par, text="方式：").grid(row=0, column=1, sticky="e")
        ttk.Combobox(frm_par, textvariable=self.kind_var, values=["process", "thread"], width=10, state="readonly").grid(row=0, column=2, sticky="w", padx=(2, 8))
        ttk.Label(frm_par, text="最大并发：").grid(row=0, column=3, sticky="e")
        ttk.Spinbox(frm_par, from_=1, to=max(64, (os.cpu_count() or 8)), textvariable=self.workers_var, width=6).grid(row=0, column=4, sticky="w", padx=(2, 8))

        # 进度 & 按钮
        frm_run = ttk.Frame(self)
        frm_run.pack(fill="x", **pad)
        self.btn_start = ttk.Button(frm_run, text="开始合并", command=self._on_start)
        self.btn_start.pack(side="left")
        self.btn_stop = ttk.Button(frm_run, text="取消任务", command=self._on_stop, state="disabled")
        self.btn_stop.pack(side="left", padx=(8, 0))
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
                self.mode_var.set(cfg.get("mode", "single"))
                self.engine_var.set(cfg.get("engine", "openpyxl"))
        except Exception:
            pass

    def _save_cfg(self):
        try:
            cfg = {
                "dir": self.dir_var.get(),
                "out": self.out_var.get(),
                "mode": self.mode_var.get(),
                "engine": self.engine_var.get(),
            }
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
            title="选择输出 Excel 文件（.xlsx）",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if f:
            self.out_var.set(f)

    def _browse_out_dir(self):
        d = filedialog.askdirectory(title="选择输出文件夹（将自动生成 merged.xlsx）")
        if d:
            self.out_var.set(d)

    def _open_out_dir(self):
        raw = self.out_var.get().strip()
        if not raw:
            return
        try:
            out_file = resolve_output_path(raw)
            out_dir = out_file.parent
            if sys.platform.startswith("win"):
                os.startfile(out_dir)  # type: ignore
            else:
                # mac/linux
                import subprocess
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(out_dir)])
        except Exception as e:
            messagebox.showwarning("提示", f"无法打开目录：{e}")

    def _on_stop(self):
        self._working = False
        self.log("收到取消请求，正在尝试终止...（当前批次完成后停止）")
        self.btn_stop.configure(state="disabled")

    def _on_start(self):
        if self._working:
            return
        in_dir = self.dir_var.get().strip()
        out_raw = self.out_var.get().strip()

        if not in_dir:
            messagebox.showwarning("提示", "请选择待合并文件夹")
            return
        if not out_raw:
            messagebox.showwarning("提示", "请选择输出路径（文件或文件夹）")
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
            dedup_cols=parse_fields(self.dedup_var.get()),
            parallel=self.parallel_var.get(),
            parallel_kind=self.kind_var.get(),
            max_workers=max(1, int(self.workers_var.get() or 1)),
            auto_width=self.auto_width_var.get(),
            engine=pick_writer_engine(self.engine_var.get()),
        )

        # engine 回退提示
        if self.engine_var.get().lower().strip() == "xlsxwriter" and opts.engine != "xlsxwriter":
            messagebox.showinfo("提示", "检测到未安装 xlsxwriter，已自动切换为 openpyxl。")

        self._save_cfg()

        # 清理 UI
        self._working = True
        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.txt.delete("1.0", "end")
        self.prog.configure(value=0, maximum=100)

        # 启动任务（用 after 让 UI 先刷新）
        self.after(50, lambda: self._run_merge(root, out_raw, opts))

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
    def _run_merge(self, root: Path, out_raw: str, opts: Options):
        try:
            files = collect_xlsx_files(root, recursive=opts.recursive)

            # 解析输出文件路径（关键修复点）
            out_file = resolve_output_path(out_raw, default_filename="merged.xlsx")

            # 排除输出文件（若位于同一目录/递归范围内）
            try:
                out_resolved = out_file.resolve()
                files = [f for f in files if f.resolve() != out_resolved]
            except Exception:
                pass

            if not files:
                self.log("未在该目录下找到 .xlsx 文件。")
                messagebox.showinfo("提示", "未找到 .xlsx 文件")
                return self._finish()

            self.log(f"发现 {len(files)} 个 .xlsx 文件，即将开始解析...")
            self.set_progress(0, maximum=len(files) + 1)

            all_pairs: List[Tuple[str, pd.DataFrame]] = []
            read_done = 0

            # ---- 读取阶段 ----
            if opts.parallel:
                max_workers = opts.max_workers
                ex_cls = ProcessPoolExecutor if opts.parallel_kind == "process" else ThreadPoolExecutor
                self.log(f"并行模式：{opts.parallel_kind}，并发数：{max_workers}")

                with ex_cls(max_workers=max_workers) as ex:
                    fut_map = {}
                    for f in files:
                        if not self._working:
                            break
                        fut = ex.submit(read_excel_file_worker, str(f), opts.drop_empty_rows, opts.add_src_file, opts.add_src_sheet)
                        fut_map[fut] = f

                    for fut in as_completed(fut_map):
                        if not self._working:
                            # 尝试取消未开始的任务（Py3.9+）
                            try:
                                ex.shutdown(wait=False, cancel_futures=True)  # type: ignore
                            except Exception:
                                pass
                            break

                        f = fut_map[fut]
                        try:
                            res = fut.result()
                        except Exception as e:
                            self.log(f"[异常] {f.name}: {e}")
                            read_done += 1
                            self.set_progress(read_done, maximum=len(files) + 1)
                            continue

                        read_done += 1
                        if res.get("ok"):
                            all_pairs.extend(res.get("pairs", []))
                            self.log(f"[读取完成] {Path(res['file']).name}（{len(res.get('pairs', []))} 个Sheet，共 {res.get('nrows', 0)} 行）")
                        else:
                            self.log(res.get("err") or f"[读取失败] {f.name}")
                        self.set_progress(read_done, maximum=len(files) + 1)
            else:
                for i, f in enumerate(files, 1):
                    if not self._working:
                        break
                    res = read_excel_file_worker(str(f), opts.drop_empty_rows, opts.add_src_file, opts.add_src_sheet)
                    if res.get("ok"):
                        all_pairs.extend(res.get("pairs", []))
                        self.log(f"[读取完成] {Path(res['file']).name}（{len(res.get('pairs', []))} 个Sheet，共 {res.get('nrows', 0)} 行）")
                    else:
                        self.log(res.get("err") or f"[读取失败] {f.name}")
                    self.set_progress(i, maximum=len(files) + 1)

            if not self._working:
                self.log("任务已取消。")
                return self._finish()

            if not all_pairs:
                self.log("未读取到任何数据，终止。")
                messagebox.showwarning("提示", "未读取到任何数据")
                return self._finish()

            # ---- 合并阶段 ----
            self.log(f"开始合并（模式：{'总表' if opts.mode == 'single' else '按Sheet'}，列集合：{'并集' if opts.union_columns else '交集'}）...")
            out_file.parent.mkdir(parents=True, exist_ok=True)

            # 写到临时文件，成功后替换
            tmp_file = out_file.with_name(out_file.stem + ".__writing__.xlsx")
            if tmp_file.exists():
                try:
                    tmp_file.unlink()
                except Exception:
                    pass

            # 如果目标文件已打开，Windows 下常见 PermissionError
            if out_file.exists():
                try:
                    # 尝试以追加方式打开一下（不写入），用于提前暴露锁定问题
                    with open(out_file, "ab"):
                        pass
                except PermissionError:
                    raise PermissionError(
                        f"输出文件正在被占用（通常是 Excel 打开了它）：\n{out_file}\n\n请关闭 Excel 后重试，或换一个输出文件名。"
                    )

            used_sheet_names: set[str] = set()

            try:
                with pd.ExcelWriter(tmp_file, engine=opts.engine) as writer:
                    if opts.mode == "single":
                        dfs = [df for (_s, df) in all_pairs]
                        merged = self._merge_dfs(dfs, opts.union_columns)

                        if opts.dedup_cols:
                            self._try_dedup(merged, opts.dedup_cols, prefix="")

                        parts = chunk_dataframe(merged, EXCEL_MAX_ROWS)
                        for idx, part in enumerate(parts, 1):
                            base_name = "merged" if len(parts) == 1 else f"merged_{idx}"
                            sheet_name = ensure_unique_sheet_name(safe_sheet_title(base_name), used_sheet_names)
                            part.to_excel(writer, sheet_name=sheet_name, index=False)
                            if opts.auto_width:
                                self._adjust_width(writer, sheet_name, part)
                            self.log(f"[写入] {sheet_name}  行数：{len(part)}")
                    else:
                        groups: Dict[str, List[pd.DataFrame]] = {}
                        for s, df in all_pairs:
                            groups.setdefault(s, []).append(df)

                        for sname, lst in groups.items():
                            merged = self._merge_dfs(lst, opts.union_columns)

                            if opts.dedup_cols:
                                self._try_dedup(merged, opts.dedup_cols, prefix=f"-{sname}")

                            parts = chunk_dataframe(merged, EXCEL_MAX_ROWS)
                            for idx, part in enumerate(parts, 1):
                                suffix = "" if len(parts) == 1 else f"_{idx}"
                                base = safe_sheet_title(sname, suffix=suffix)
                                sheet_name = ensure_unique_sheet_name(base, used_sheet_names)
                                part.to_excel(writer, sheet_name=sheet_name, index=False)
                                if opts.auto_width:
                                    self._adjust_width(writer, sheet_name, part)
                                self.log(f"[写入] {sheet_name}  行数：{len(part)}")
            except PermissionError as e:
                # 写临时文件/替换也可能 PermissionError（目录没权限、文件被占用等）
                raise PermissionError(
                    f"{e}\n\n常见原因：\n"
                    f"1) 你选择的输出路径是“文件夹”，但程序无法写入该目录（权限/只读/被占用）。\n"
                    f"2) 目标文件正在被 Excel 打开（请关闭 Excel）。\n"
                    f"3) 输出路径被杀毒/网盘同步锁定。\n\n"
                    f"当前解析到的输出文件路径是：\n{out_file}"
                )

            # 临时文件成功写完 -> 替换正式文件
            atomic_replace(tmp_file, out_file)

            self.set_progress(len(files) + 1, maximum=len(files) + 1)
            self.log(f"✅ 合并完成，输出：{out_file}")

            # 完成后打开目录（Windows）
            try:
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

    def _merge_dfs(self, dfs: List[pd.DataFrame], union_columns: bool) -> pd.DataFrame:
        if not dfs:
            return pd.DataFrame()

        if union_columns:
            return pd.concat(dfs, ignore_index=True, sort=False)

        # 交集列：保持第一张表的列顺序
        commons = set(dfs[0].columns)
        for d in dfs[1:]:
            commons &= set(d.columns)
        ordered = [c for c in dfs[0].columns if c in commons]
        if not ordered:
            return pd.DataFrame()
        return pd.concat([d[ordered].copy() for d in dfs], ignore_index=True)

    def _try_dedup(self, df: pd.DataFrame, cols: List[str], prefix: str = "") -> None:
        try:
            df.drop_duplicates(subset=cols, inplace=True, ignore_index=True)
        except Exception as e:
            self.log(f"[去重警告{prefix}] {e}")

    def _finish(self):
        self._working = False
        self.btn_start.configure(state="normal")
        self.btn_stop.configure(state="disabled")

    def _adjust_width(self, writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame, sample_rows: int = 200):
        """基于前 sample_rows 行估算列宽（openpyxl/xlsxwriter）"""
        try:
            if writer.engine == "openpyxl":
                from openpyxl.utils import get_column_letter
                ws = writer.sheets.get(sheet_name)
                if ws is None and hasattr(writer, "book"):
                    ws = writer.book[sheet_name]
                if ws is None:
                    return
                for i, col in enumerate(df.columns, start=1):
                    series = df[col].astype(str).head(sample_rows)
                    max_len = max([len(str(col))] + [len(v) for v in series.tolist()]) if len(series) else len(str(col))
                    max_len = min(60, max_len + 2)
                    ws.column_dimensions[get_column_letter(i)].width = max_len
                ws.freeze_panes = "A2"
            elif writer.engine == "xlsxwriter":
                ws = writer.sheets.get(sheet_name)
                if ws is None:
                    return
                for i, col in enumerate(df.columns):
                    series = df[col].astype(str).head(sample_rows)
                    max_len = max([len(str(col))] + [len(v) for v in series.tolist()]) if len(series) else len(str(col))
                    max_len = min(60, max_len + 2)
                    ws.set_column(i, i, max_len)
                ws.freeze_panes(1, 0)
        except Exception:
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


