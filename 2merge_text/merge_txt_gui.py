# -*- coding: utf-8 -*-
"""
merge_txt_gui.py

一个用于批量合并文件夹中 .txt 文本的图形界面工具（Tkinter）。
特性：
- 选择输入/输出文件夹
- 每 N 个 .txt 合并为 1 个
- 支持并发读取（可控线程数）
- 标题插入（文件名或首个非空行），可避免重复标题
- 文档间统一空行数 & 可选分隔符行
- 自然排序（数字识别），保证写出顺序稳定
- 实时进度条与滚动日志
- 多编码回退读取（utf-8/gbk/gb18030/big5/utf-16/utf-32）

仅依赖 Python 标准库。
"""
from __future__ import annotations

import os
import math
import re
import threading
import queue
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText


# -------------------------- 通用工具函数 --------------------------

def natural_key(s: str):
    """
    按“自然顺序”拆分：把数字片段转为整数，其它部分转小写字符串。
    例如：file2 < file10
    """
    return [int(p) if p.isdigit() else p.lower() for p in re.split(r'(\d+)', str(s))]


def list_txt_files(input_dir: Path):
    """列出文件夹内所有 .txt（不递归），按自然顺序排序。"""
    files = [p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() == '.txt']
    files.sort(key=lambda p: natural_key(p.name))
    return files


def read_text_with_fallback(path: Path, encodings=None):
    """
    用多编码回退方式读取文本。
    返回 (text, used_encoding)
    """
    if encodings is None:
        encodings = ['utf-8-sig', 'utf-8', 'gb18030', 'gbk', 'big5', 'utf-16', 'utf-32']
    last_err = None
    for enc in encodings:
        try:
            with path.open('r', encoding=enc) as f:
                return f.read(), enc
        except Exception as e:
            last_err = e
    # 最后兜底：二进制读 + 忽略错误
    try:
        data = path.read_bytes()
        return data.decode('utf-8', errors='ignore'), 'utf-8(ignore)'
    except Exception:
        raise last_err or RuntimeError(f"无法读取：{path}")


def derive_title_and_body(path: Path, raw_content: str, *,
                          title_mode: str = 'filename',
                          avoid_duplicate: bool = True):
    """
    根据 title_mode 生成标题，并返回 (title_line, body_content_without_dup_title)
    title_mode: 'filename' | 'firstline'
    avoid_duplicate: 若内容首行（去空白）与标题相同，则不再重复插入该标题
    """
    title = path.stem
    if title_mode == 'firstline':
        # 找到首个非空行作为标题
        for line in raw_content.splitlines():
            if line.strip():
                title = line.strip()
                break

    body = raw_content
    if avoid_duplicate:
        # 如果正文以标题开头（包含可能的 BOM/空白），避免重复
        norm_body = body.lstrip('\ufeff').lstrip()
        if norm_body.startswith(title):
            return title, body
    return title, body


def build_group_text(file_paths, content_map, *,
                     title_mode='filename',
                     avoid_duplicate_title=True,
                     blank_lines=3,
                     use_separator=False,
                     separator_line=''):
    """
    构建一组（若干 .txt）合并后的完整文本。
    每篇结构：<标题>\n<原始内容>\n\n...（若干空行 + 可选分隔符）
    """
    parts = []
    for i, p in enumerate(file_paths):
        raw = content_map[p]['text']
        title, body = derive_title_and_body(
            p, raw,
            title_mode=title_mode,
            avoid_duplicate=avoid_duplicate_title
        )

        # 标题置顶一行；正文保持原样；末尾附加统一间隔
        parts.append(f"{title}\n{body.rstrip()}\n")  # 去掉末尾多余空白行，再补一个换行
        if use_separator and separator_line.strip():
            parts.append(separator_line.rstrip() + "\n")

        parts.append("\n" * blank_lines)

    # 合并后整体去除最末尾多余空白，避免文件末尾无限空行
    merged = "".join(parts).rstrip() + "\n"
    return merged


# -------------------------- GUI 主应用 --------------------------

class MergeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TXT 批量合并器（支持标题/空行间隔/并发/进度展示）")
        self.minsize(920, 560)

        # 绑定变量
        self.input_dir_var = tk.StringVar(value="")
        self.output_dir_var = tk.StringVar(value="")
        self.batch_size_var = tk.IntVar(value=10)
        self.thread_var = tk.IntVar(value=min(8, (os.cpu_count() or 4)))
        self.blank_lines_var = tk.IntVar(value=3)
        self.use_separator_var = tk.BooleanVar(value=False)
        self.separator_var = tk.StringVar(value="")  # 例如 "——— 分隔线 ———"
        self.title_mode_var = tk.StringVar(value="filename")  # 'filename' or 'firstline'
        self.avoid_dup_title_var = tk.BooleanVar(value=True)
        self.output_encoding_var = tk.StringVar(value="utf-8")  # 输出编码
        self.sort_mode_var = tk.StringVar(value="natural")  # 'natural' | 'lex' | 'mtime'

        # 控制状态
        self._worker_thread: threading.Thread | None = None
        self._stop_event = threading.Event()
        self._ui_queue = queue.Queue()
        self._progress_total = 100
        self._progress_val = 0

        self._build_ui()
        self._schedule_ui_pump()

    # ---------- UI 构建 ----------
    def _build_ui(self):
        pad = {'padx': 8, 'pady': 6}

        # 路径行
        path_frame = ttk.LabelFrame(self, text="路径")
        path_frame.pack(fill='x', **pad)

        ttk.Label(path_frame, text="输入文件夹：").grid(row=0, column=0, sticky='w')
        ttk.Entry(path_frame, textvariable=self.input_dir_var, width=70).grid(row=0, column=1, sticky='we', padx=6)
        ttk.Button(path_frame, text="浏览…", command=self.select_input_dir).grid(row=0, column=2)

        ttk.Label(path_frame, text="输出文件夹：").grid(row=1, column=0, sticky='w')
        ttk.Entry(path_frame, textvariable=self.output_dir_var, width=70).grid(row=1, column=1, sticky='we', padx=6)
        ttk.Button(path_frame, text="浏览…", command=self.select_output_dir).grid(row=1, column=2)

        path_frame.columnconfigure(1, weight=1)

        # 参数行
        opts = ttk.LabelFrame(self, text="参数")
        opts.pack(fill='x', **pad)

        # 左半：数量/线程/间隔
        left = ttk.Frame(opts)
        left.grid(row=0, column=0, sticky='w', padx=4)
        ttk.Label(left, text="每份合并数量：").grid(row=0, column=0, sticky='w')
        ttk.Spinbox(left, from_=1, to=99999, textvariable=self.batch_size_var, width=8).grid(row=0, column=1, sticky='w', padx=(0, 12))

        ttk.Label(left, text="线程数：").grid(row=0, column=2, sticky='w')
        ttk.Spinbox(left, from_=1, to=max(64, (os.cpu_count() or 4)), textvariable=self.thread_var, width=8).grid(row=0, column=3, sticky='w', padx=(0, 12))

        ttk.Label(left, text="文档间空行数：").grid(row=0, column=4, sticky='w')
        ttk.Spinbox(left, from_=0, to=50, textvariable=self.blank_lines_var, width=8).grid(row=0, column=5, sticky='w')

        # 中间：标题与排序
        mid = ttk.Frame(opts)
        mid.grid(row=0, column=1, sticky='w', padx=12)

        title_mode_frame = ttk.LabelFrame(mid, text="标题来源")
        title_mode_frame.grid(row=0, column=0, sticky='w', padx=(0, 12))
        ttk.Radiobutton(title_mode_frame, text="文件名", value="filename", variable=self.title_mode_var).grid(row=0, column=0, sticky='w')
        ttk.Radiobutton(title_mode_frame, text="文件首个非空行", value="firstline", variable=self.title_mode_var).grid(row=0, column=1, sticky='w')
        ttk.Checkbutton(title_mode_frame, text="避免重复标题", variable=self.avoid_dup_title_var).grid(row=1, column=0, columnspan=2, sticky='w')

        sort_frame = ttk.LabelFrame(mid, text="排序方式")
        sort_frame.grid(row=0, column=1, sticky='w')
        ttk.Radiobutton(sort_frame, text="自然顺序(默认)", value="natural", variable=self.sort_mode_var).grid(row=0, column=0, sticky='w')
        ttk.Radiobutton(sort_frame, text="字典序", value="lex", variable=self.sort_mode_var).grid(row=0, column=1, sticky='w')
        ttk.Radiobutton(sort_frame, text="按修改时间", value="mtime", variable=self.sort_mode_var).grid(row=0, column=2, sticky='w')

        # 右侧：分隔符与输出编码
        right = ttk.Frame(opts)
        right.grid(row=0, column=2, sticky='w', padx=12)
        ttk.Checkbutton(right, text="在每篇后插入分隔符行", variable=self.use_separator_var).grid(row=0, column=0, sticky='w')
        ttk.Entry(right, textvariable=self.separator_var, width=30).grid(row=0, column=1, sticky='w', padx=(6, 0))
        ttk.Label(right, text="输出编码：").grid(row=1, column=0, sticky='e', pady=(6, 0))
        ttk.Combobox(right, textvariable=self.output_encoding_var, values=["utf-8", "utf-8-sig", "gbk", "gb18030"], width=12, state="readonly").grid(row=1, column=1, sticky='w', pady=(6, 0))

        # 操作按钮
        ops = ttk.Frame(self)
        ops.pack(fill='x', **pad)
        self.start_btn = ttk.Button(ops, text="开始合并", command=self.start_merge)
        self.start_btn.pack(side='left')
        self.stop_btn = ttk.Button(ops, text="取消", command=self.stop_merge, state='disabled')
        self.stop_btn.pack(side='left', padx=(8, 0))
        self.open_out_btn = ttk.Button(ops, text="打开输出文件夹", command=self.open_output_dir)
        self.open_out_btn.pack(side='left', padx=(8, 0))

        # 进度 & 状态
        prog = ttk.Frame(self)
        prog.pack(fill='x', **pad)
        self.progress = ttk.Progressbar(prog, mode='determinate', maximum=100)
        self.progress.pack(fill='x')
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(prog, textvariable=self.status_var, anchor='w').pack(fill='x')

        # 日志
        log_frame = ttk.LabelFrame(self, text="实时日志")
        log_frame.pack(fill='both', expand=True, **pad)
        self.log_widget = ScrolledText(log_frame, height=16)
        self.log_widget.pack(fill='both', expand=True)

    # ---------- 事件处理 ----------
    def select_input_dir(self):
        d = filedialog.askdirectory(title="选择包含 .txt 的输入文件夹")
        if d:
            self.input_dir_var.set(d)
            # 如果还未设置输出文件夹，默认使用输入目录下的 merged_output
            if not self.output_dir_var.get():
                self.output_dir_var.set(str(Path(d) / "merged_output"))

    def select_output_dir(self):
        d = filedialog.askdirectory(title="选择输出文件夹")
        if d:
            self.output_dir_var.set(d)

    def open_output_dir(self):
        p = Path(self.output_dir_var.get().strip())
        if not p.exists():
            messagebox.showinfo("提示", "输出文件夹不存在。")
            return
        try:
            if os.name == 'nt':
                os.startfile(str(p))
            elif os.name == 'posix':
                import subprocess
                subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', str(p)])
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件夹：{e}")

    def log(self, msg: str):
        self._ui_queue.put(("log", msg))

    def set_status(self, msg: str):
        self._ui_queue.put(("status", msg))

    def set_progress_total(self, total: int):
        self._progress_total = max(1, total)
        self._ui_queue.put(("progress_config", self._progress_total))

    def add_progress(self, delta: int = 1):
        self._ui_queue.put(("progress_add", delta))

    def _schedule_ui_pump(self):
        """定时从队列取事件刷新 UI（保证线程安全）。"""
        try:
            while True:
                kind, payload = self._ui_queue.get_nowait()
                if kind == "log":
                    self.log_widget.insert('end', f"[{datetime.now().strftime('%H:%M:%S')}] {payload}\n")
                    self.log_widget.see('end')
                elif kind == "status":
                    self.status_var.set(payload)
                elif kind == "progress_config":
                    self.progress.configure(maximum=int(payload))
                    self._progress_val = 0
                    self.progress['value'] = 0
                elif kind == "progress_add":
                    self._progress_val += int(payload)
                    self.progress['value'] = min(self._progress_val, self._progress_total)
        except queue.Empty:
            pass
        # 100ms 刷新一次
        self.after(100, self._schedule_ui_pump)

    def start_merge(self):
        try:
            input_dir = Path(self.input_dir_var.get().strip())
            output_dir = Path(self.output_dir_var.get().strip())
            if not input_dir.exists() or not input_dir.is_dir():
                messagebox.showerror("错误", "请输入有效的输入文件夹。")
                return
            if not output_dir.exists():
                output_dir.mkdir(parents=True, exist_ok=True)

            batch_size = max(1, int(self.batch_size_var.get()))
            threads = max(1, int(self.thread_var.get()))
            blank_lines = max(0, int(self.blank_lines_var.get()))
            use_separator = bool(self.use_separator_var.get())
            separator = self.separator_var.get()
            title_mode = self.title_mode_var.get()
            avoid_dup = bool(self.avoid_dup_title_var.get())
            out_encoding = self.output_encoding_var.get()
            sort_mode = self.sort_mode_var.get()

            # 切换按钮状态
            self.start_btn.config(state='disabled')
            self.stop_btn.config(state='normal')
            self.log("开始任务 …")
            self.set_status("初始化 …")

            self._stop_event.clear()
            self._worker_thread = threading.Thread(
                target=self._do_merge_work,
                kwargs=dict(
                    input_dir=input_dir,
                    output_dir=output_dir,
                    batch_size=batch_size,
                    threads=threads,
                    blank_lines=blank_lines,
                    use_separator=use_separator,
                    separator=separator,
                    title_mode=title_mode,
                    avoid_dup=avoid_dup,
                    out_encoding=out_encoding,
                    sort_mode=sort_mode,
                ),
                daemon=True
            )
            self._worker_thread.start()

        except Exception as e:
            messagebox.showerror("错误", f"启动失败：{e}")

    def stop_merge(self):
        if self._worker_thread and self._worker_thread.is_alive():
            self._stop_event.set()
            self.log("收到取消请求，正在停止…")
            self.set_status("取消中 …")

    # ---------- 核心工作流 ----------
    def _do_merge_work(self, *, input_dir: Path, output_dir: Path, batch_size: int,
                       threads: int, blank_lines: int, use_separator: bool,
                       separator: str, title_mode: str, avoid_dup: bool,
                       out_encoding: str, sort_mode: str):
        try:
            # 列出文件
            files = list_txt_files(input_dir)
            if sort_mode == 'lex':
                files.sort(key=lambda p: p.name.lower())
            elif sort_mode == 'mtime':
                files.sort(key=lambda p: p.stat().st_mtime)

            total = len(files)
            if total == 0:
                self.log("输入文件夹中未找到 .txt 文件。")
                self.set_status("空目录")
                return

            self.log(f"共发现 {total} 个 .txt 文件。")
            num_groups = math.ceil(total / batch_size)
            # 进度：读取 total 次 + 写组 num_groups 次
            self.set_progress_total(total + num_groups)

            # 并发读取
            self.set_status("并发读取文件 …")
            content_map = {}  # {Path: {'text': str, 'enc': str}}
            read_errors = []

            def read_one(p: Path):
                if self._stop_event.is_set():
                    return None
                self.log(f"读取：{p.name}")
                txt, enc = read_text_with_fallback(p)
                return p, txt, enc

            with ThreadPoolExecutor(max_workers=threads) as ex:
                futures = {ex.submit(read_one, p): p for p in files}
                for fut in as_completed(futures):
                    if self._stop_event.is_set():
                        break
                    p = futures[fut]
                    try:
                        res = fut.result()
                        if res is None:
                            continue
                        rp, txt, enc = res
                        content_map[rp] = {'text': txt, 'enc': enc}
                        self.add_progress(1)
                        self.log(f"完成：{rp.name}（编码：{enc}）")
                    except Exception as e:
                        read_errors.append((p, str(e)))
                        self.add_progress(1)
                        self.log(f"读取失败：{p.name} | {e}")

            if self._stop_event.is_set():
                self.set_status("已取消")
                self.log("任务已取消（读取阶段）。")
                return

            if read_errors:
                self.log(f"共有 {len(read_errors)} 个文件读取失败，已跳过。")

            # 分组写出
            self.set_status("开始写出合并文件 …")
            for gi in range(num_groups):
                if self._stop_event.is_set():
                    break
                start = gi * batch_size
                group_files = [p for p in files[start:start + batch_size] if p in content_map]
                if not group_files:
                    # 这组全失败/为空
                    self.add_progress(1)
                    continue

                start_idx = start + 1
                end_idx = start + len(group_files)
                out_name = f"group_{gi+1:03d}_{start_idx}-{end_idx}.txt"
                out_path = output_dir / out_name

                self.set_status(f"写出第 {gi+1}/{num_groups} 组：{start_idx}-{end_idx}")
                self.log(f"合并组：{start_idx}-{end_idx} → {out_name}")

                merged_text = build_group_text(
                    group_files, content_map,
                    title_mode=title_mode,
                    avoid_duplicate_title=avoid_dup,
                    blank_lines=blank_lines,
                    use_separator=use_separator,
                    separator_line=separator
                )

                out_path.write_text(merged_text, encoding=out_encoding, newline='\n')
                self.add_progress(1)
                self.log(f"写出完成：{out_name}")

            if self._stop_event.is_set():
                self.set_status("已取消")
                self.log("任务已取消（写出阶段）。")
                return

            self.set_status("全部完成 ✔")
            self.log("✅ 任务完成！你可以点击“打开输出文件夹”查看结果。")

        except Exception as e:
            self.set_status("出错")
            self.log(f"❌ 发生错误：{e}")
            messagebox.showerror("错误", f"任务执行失败：\n{e}")
        finally:
            # 恢复按钮
            self.start_btn.config(state='normal')
            self.stop_btn.config(state='disabled')


if __name__ == "__main__":
    app = MergeApp()
    app.mainloop()
