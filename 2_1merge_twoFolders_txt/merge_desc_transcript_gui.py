#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Merge two folders of txt files by shared episode index like 【P1】, 【P2】...

Enhancements:
- Merge only a range of P indices (e.g., P10~P80)
- Button to open output directory
- Record unmatched files and export logs to .txt

Use case:
- Folder A: Bilibili video descriptions (short text)
- Folder B: Video transcripts (longer text)
- Filenames may differ, but the index token like 【P12】 matches.

Features:
- Tkinter GUI: choose desc folder, transcript folder, output folder
- Match by episode index extracted from filename: 【P<number>】 or [P<number>]
- Output name uses the more "complete" (longer) filename among the pair
- Options: order (desc first / transcript first), add section headers, recursive scan, overwrite
- Robust decoding: tries utf-8-sig/utf-8/gb18030/gbk
"""

from __future__ import annotations

import os
import re
import sys
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText


# -----------------------
# Core logic
# -----------------------

INDEX_RE = re.compile(r"[【\[]\s*P\s*(\d+)\s*[】\]]", re.IGNORECASE)


def extract_index(name: str) -> Optional[int]:
    """Extract episode index from filename, e.g. 【P12】xxx.txt or [P12]xxx.txt"""
    m = INDEX_RE.search(name)
    if not m:
        return None
    try:
        return int(m.group(1))
    except ValueError:
        return None


def sanitize_filename(name: str) -> str:
    """Sanitize filename for Windows."""
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip()


def choose_most_complete(paths: List[Path]) -> Path:
    """
    Choose the 'most complete' file among duplicates.
    Primary: longest stem length; Secondary: larger file size.
    """
    def key(p: Path):
        try:
            size = p.stat().st_size
        except OSError:
            size = 0
        return (len(p.stem), size)

    return sorted(paths, key=key, reverse=True)[0]


def read_text_with_fallback(path: Path) -> str:
    """Read text with encoding fallbacks (Windows-friendly)."""
    data = path.read_bytes()
    for enc in ("utf-8-sig", "utf-8", "gb18030", "gbk"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def safe_output_path(out_dir: Path, filename: str, overwrite: bool) -> Path:
    """Return a non-conflicting output path if not overwriting."""
    out_dir.mkdir(parents=True, exist_ok=True)

    base = sanitize_filename(filename)
    if not base.lower().endswith(".txt"):
        base += ".txt"

    candidate = out_dir / base
    if overwrite or not candidate.exists():
        return candidate

    stem = candidate.stem
    suffix = candidate.suffix
    for i in range(1, 10_000):
        c = out_dir / f"{stem}_{i}{suffix}"
        if not c.exists():
            return c

    raise RuntimeError("Too many conflicting output files.")


@dataclass
class ScanResult:
    mapping: Dict[int, List[Path]]
    no_index_files: List[Path]
    duplicates: Dict[int, List[Path]]


def scan_txt_folder(folder: Path, recursive: bool) -> ScanResult:
    """Scan txt files and build index -> list[Path] mapping."""
    pattern = "**/*.txt" if recursive else "*.txt"
    files = list(folder.glob(pattern))

    mapping: Dict[int, List[Path]] = {}
    no_index: List[Path] = []
    duplicates: Dict[int, List[Path]] = {}

    for f in files:
        if not f.is_file():
            continue
        idx = extract_index(f.name)
        if idx is None:
            no_index.append(f)
            continue
        mapping.setdefault(idx, []).append(f)

    for idx, lst in mapping.items():
        if len(lst) > 1:
            duplicates[idx] = lst

    return ScanResult(mapping=mapping, no_index_files=no_index, duplicates=duplicates)


def merge_pair(
    desc_path: Path,
    transcript_path: Path,
    order: str,
    add_headers: bool,
) -> str:
    """Merge two files into one string."""
    desc = read_text_with_fallback(desc_path).strip()
    trans = read_text_with_fallback(transcript_path).strip()

    sep = "\n\n"
    if add_headers:
        part_desc = f"===== 视频简介（B站）=====\n{desc}\n===== 简介结束 ====="
        part_trans = f"===== 视频转写（文本）=====\n{trans}\n===== 转写结束 ====="
    else:
        part_desc = desc
        part_trans = trans

    if order == "desc_first":
        return part_desc + sep + part_trans + "\n"
    else:
        return part_trans + sep + part_desc + "\n"


def open_folder(path: Path) -> None:
    """Open folder in file explorer."""
    path = path.resolve()
    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)

    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception as e:
        raise RuntimeError(f"打开目录失败：{e}")


def parse_int_optional(s: str) -> Optional[int]:
    s = s.strip()
    if not s:
        return None
    if not re.fullmatch(r"\d+", s):
        return None
    return int(s)


def filter_indices(indices: List[int], p_min: Optional[int], p_max: Optional[int]) -> List[int]:
    out = []
    for x in indices:
        if p_min is not None and x < p_min:
            continue
        if p_max is not None and x > p_max:
            continue
        out.append(x)
    return out


# -----------------------
# GUI
# -----------------------

class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("按【P序号】合并简介 + 转写文本（TXT）")
        self.geometry("1020x720")
        self.minsize(930, 640)

        self.desc_dir = tk.StringVar()
        self.trans_dir = tk.StringVar()
        self.out_dir = tk.StringVar()

        self.recursive = tk.BooleanVar(value=False)
        self.overwrite = tk.BooleanVar(value=False)
        self.add_headers = tk.BooleanVar(value=True)

        self.order = tk.StringVar(value="desc_first")  # desc_first / transcript_first

        # Range: allow empty
        self.p_min_str = tk.StringVar(value="")
        self.p_max_str = tk.StringVar(value="")

        self._build_ui()

        self._last_scan: Optional[Tuple[ScanResult, ScanResult]] = None
        self._last_common: List[int] = []
        self._last_only_desc: List[int] = []
        self._last_only_trans: List[int] = []

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True)

        # Folder selectors
        row = 0
        ttk.Label(frm, text="简介文件夹（bili_desc_txt）：").grid(row=row, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.desc_dir).grid(row=row, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="选择...", command=self._pick_desc).grid(row=row, column=2, sticky="ew", **pad)

        row += 1
        ttk.Label(frm, text="转写文件夹（yuebao_text）：").grid(row=row, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.trans_dir).grid(row=row, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="选择...", command=self._pick_trans).grid(row=row, column=2, sticky="ew", **pad)

        row += 1
        ttk.Label(frm, text="输出文件夹（可不选，默认自动创建 merged_output）：").grid(row=row, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.out_dir).grid(row=row, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="选择...", command=self._pick_out).grid(row=row, column=2, sticky="ew", **pad)

        frm.columnconfigure(1, weight=1)

        # Options
        row += 1
        opt = ttk.LabelFrame(frm, text="选项")
        opt.grid(row=row, column=0, columnspan=3, sticky="ew", padx=10, pady=10)

        ttk.Checkbutton(opt, text="递归扫描子目录", variable=self.recursive).grid(row=0, column=0, sticky="w", padx=10, pady=6)
        ttk.Checkbutton(opt, text="覆盖输出（同名直接覆盖）", variable=self.overwrite).grid(row=0, column=1, sticky="w", padx=10, pady=6)
        ttk.Checkbutton(opt, text="合并时加入“简介/转写”标题分隔", variable=self.add_headers).grid(row=0, column=2, sticky="w", padx=10, pady=6)

        ttk.Label(opt, text="合并顺序：").grid(row=1, column=0, sticky="w", padx=10, pady=6)
        ttk.Radiobutton(opt, text="简介在前", value="desc_first", variable=self.order).grid(row=1, column=1, sticky="w", padx=10, pady=6)
        ttk.Radiobutton(opt, text="转写在前", value="transcript_first", variable=self.order).grid(row=1, column=2, sticky="w", padx=10, pady=6)

        # Range
        rng = ttk.LabelFrame(frm, text="只合并某个 P 序号范围（可选）")
        row += 1
        rng.grid(row=row, column=0, columnspan=3, sticky="ew", padx=10, pady=6)

        ttk.Label(rng, text="从 P").grid(row=0, column=0, sticky="w", padx=10, pady=8)
        e1 = ttk.Entry(rng, textvariable=self.p_min_str, width=10)
        e1.grid(row=0, column=1, sticky="w", padx=4, pady=8)
        ttk.Label(rng, text="到 P").grid(row=0, column=2, sticky="w", padx=10, pady=8)
        e2 = ttk.Entry(rng, textvariable=self.p_max_str, width=10)
        e2.grid(row=0, column=3, sticky="w", padx=4, pady=8)
        ttk.Label(rng, text="（留空=不限制；例如 10 和 80 表示只处理 P10~P80）").grid(row=0, column=4, sticky="w", padx=10, pady=8)

        # Actions
        row += 1
        act = ttk.Frame(frm)
        act.grid(row=row, column=0, columnspan=3, sticky="ew", padx=10, pady=6)

        ttk.Button(act, text="1) 扫描并匹配", command=self._scan).pack(side="left")
        ttk.Button(act, text="2) 开始合并", command=self._merge).pack(side="left", padx=10)
        ttk.Button(act, text="打开输出目录", command=self._open_out_dir).pack(side="left", padx=10)
        ttk.Button(act, text="导出日志", command=self._export_log).pack(side="left", padx=10)
        ttk.Button(act, text="清空日志", command=self._clear_log).pack(side="left", padx=10)

        # Match list + Log
        row += 1
        mid = ttk.Panedwindow(frm, orient="horizontal")
        mid.grid(row=row, column=0, columnspan=3, sticky="nsew", padx=10, pady=10)
        frm.rowconfigure(row, weight=1)

        # Left: match tree
        left = ttk.Frame(mid)
        mid.add(left, weight=1)

        ttk.Label(left, text="匹配结果（按序号）：").pack(anchor="w")

        cols = ("idx", "desc", "trans", "outname")
        self.tree = ttk.Treeview(left, columns=cols, show="headings", height=18)
        self.tree.heading("idx", text="P序号")
        self.tree.heading("desc", text="简介文件")
        self.tree.heading("trans", text="转写文件")
        self.tree.heading("outname", text="输出文件名（更完整者）")

        self.tree.column("idx", width=70, anchor="center")
        self.tree.column("desc", width=290, anchor="w")
        self.tree.column("trans", width=290, anchor="w")
        self.tree.column("outname", width=320, anchor="w")

        self.tree.pack(fill="both", expand=True)

        # Right: log
        right = ttk.Frame(mid)
        mid.add(right, weight=1)

        ttk.Label(right, text="日志：").pack(anchor="w")
        self.log = ScrolledText(right, wrap="word")
        self.log.pack(fill="both", expand=True)

        self._log("提示：先点【扫描并匹配】，确认匹配数量无误后再点【开始合并】。\n"
                  "新增：可填写 P 范围、可一键打开输出目录、可导出日志。\n")

    def _pick_desc(self) -> None:
        p = filedialog.askdirectory(title="选择简介文件夹（bili_desc_txt）")
        if p:
            self.desc_dir.set(p)

    def _pick_trans(self) -> None:
        p = filedialog.askdirectory(title="选择转写文件夹（yuebao_text）")
        if p:
            self.trans_dir.set(p)

    def _pick_out(self) -> None:
        p = filedialog.askdirectory(title="选择输出文件夹（输出合并后的 txt）")
        if p:
            self.out_dir.set(p)

    def _clear_log(self) -> None:
        self.log.delete("1.0", "end")

    def _log(self, msg: str) -> None:
        self.log.insert("end", msg + "\n")
        self.log.see("end")

    def _get_range(self) -> Tuple[Optional[int], Optional[int]]:
        pmin = parse_int_optional(self.p_min_str.get())
        pmax = parse_int_optional(self.p_max_str.get())

        # If user typed non-digit, treat as error
        if self.p_min_str.get().strip() and pmin is None:
            raise ValueError("P范围的“从 P”必须是整数或留空。")
        if self.p_max_str.get().strip() and pmax is None:
            raise ValueError("P范围的“到 P”必须是整数或留空。")

        if pmin is not None and pmax is not None and pmin > pmax:
            raise ValueError("P范围不合法：从 P 不能大于 到 P。")

        return pmin, pmax

    def _validate_dirs(self) -> Optional[Tuple[Path, Path, Path]]:
        d = self.desc_dir.get().strip()
        t = self.trans_dir.get().strip()
        o = self.out_dir.get().strip()

        if not d or not t:
            messagebox.showerror("缺少路径", "请先选择【简介文件夹】和【转写文件夹】。")
            return None

        desc = Path(d)
        trans = Path(t)

        if not desc.exists() or not desc.is_dir():
            messagebox.showerror("路径错误", f"简介文件夹不存在或不是目录：\n{desc}")
            return None
        if not trans.exists() or not trans.is_dir():
            messagebox.showerror("路径错误", f"转写文件夹不存在或不是目录：\n{trans}")
            return None

        if o:
            out = Path(o)
        else:
            out = desc.parent / "merged_output"

        return desc, trans, out

    def _scan(self) -> None:
        v = self._validate_dirs()
        if not v:
            return

        try:
            pmin, pmax = self._get_range()
        except ValueError as e:
            messagebox.showerror("范围输入错误", str(e))
            return

        desc_dir, trans_dir, out_dir = v
        recursive = self.recursive.get()

        self._log("开始扫描...")
        self._log(f"简介目录：{desc_dir}")
        self._log(f"转写目录：{trans_dir}")
        self._log(f"输出目录：{out_dir}")
        self._log(f"递归扫描：{recursive}")
        self._log(f"P范围：{('全部' if (pmin is None and pmax is None) else f'P{pmin or ''} ~ P{pmax or ''}')}")

        desc_scan = scan_txt_folder(desc_dir, recursive=recursive)
        trans_scan = scan_txt_folder(trans_dir, recursive=recursive)
        self._last_scan = (desc_scan, trans_scan)

        # Clear tree
        for item in self.tree.get_children():
            self.tree.delete(item)

        desc_ids_all = sorted(desc_scan.mapping.keys())
        trans_ids_all = sorted(trans_scan.mapping.keys())

        # Apply range filtering for matching/summary
        desc_ids = set(filter_indices(desc_ids_all, pmin, pmax))
        trans_ids = set(filter_indices(trans_ids_all, pmin, pmax))
        common = sorted(desc_ids & trans_ids)
        only_desc = sorted(desc_ids - trans_ids)
        only_trans = sorted(trans_ids - desc_ids)

        self._last_common = common
        self._last_only_desc = only_desc
        self._last_only_trans = only_trans

        self._log(f"\n扫描完成（已应用 P 范围过滤后统计）：")
        self._log(f"- 简介序号数量：{len(desc_ids)}")
        self._log(f"- 转写序号数量：{len(trans_ids)}")
        self._log(f"- 成功匹配的序号数量：{len(common)}")
        self._log(f"- 仅简介存在的序号数量：{len(only_desc)}")
        self._log(f"- 仅转写存在的序号数量：{len(only_trans)}")

        # Additional diagnostics (full-scan, not range-limited)
        if desc_scan.no_index_files:
            self._log(f"\n[提示] 简介目录中有 {len(desc_scan.no_index_files)} 个文件无法解析【P序号】，已忽略（可导出日志查看列表）。")
        if trans_scan.no_index_files:
            self._log(f"\n[提示] 转写目录中有 {len(trans_scan.no_index_files)} 个文件无法解析【P序号】，已忽略（可导出日志查看列表）。")

        if desc_scan.duplicates:
            dups = sorted(desc_scan.duplicates.keys())
            self._log(f"\n[注意] 简介目录存在重复序号（将自动选择“更完整”的那个作为合并对象）：{dups[:30]}{'...' if len(dups)>30 else ''}")
        if trans_scan.duplicates:
            dups = sorted(trans_scan.duplicates.keys())
            self._log(f"\n[注意] 转写目录存在重复序号（将自动选择“更完整”的那个作为合并对象）：{dups[:30]}{'...' if len(dups)>30 else ''}")

        # Fill tree with matched results
        for idx in common:
            desc_path = choose_most_complete(desc_scan.mapping[idx])
            trans_path = choose_most_complete(trans_scan.mapping[idx])

            out_name = trans_path.name if len(trans_path.stem) >= len(desc_path.stem) else desc_path.name

            self.tree.insert(
                "",
                "end",
                values=(f"P{idx}", desc_path.name, trans_path.name, out_name),
            )

        # Record unmatched details into log for export
        if only_desc:
            self._log(f"\n仅简介存在的序号（前 60 个）：{only_desc[:60]}{'...' if len(only_desc)>60 else ''}")
            self._log("（提示：这些序号在转写目录未找到对应文件）")
        if only_trans:
            self._log(f"\n仅转写存在的序号（前 60 个）：{only_trans[:60]}{'...' if len(only_trans)>60 else ''}")
            self._log("（提示：这些序号在简介目录未找到对应文件）")

        # List no-index files (names only, avoid too long but export will keep full)
        if desc_scan.no_index_files:
            names = [p.name for p in desc_scan.no_index_files[:50]]
            self._log(f"\n简介目录无法解析 P 序号的文件（前 50 个）：{names}{'...' if len(desc_scan.no_index_files)>50 else ''}")
        if trans_scan.no_index_files:
            names = [p.name for p in trans_scan.no_index_files[:50]]
            self._log(f"\n转写目录无法解析 P 序号的文件（前 50 个）：{names}{'...' if len(trans_scan.no_index_files)>50 else ''}")

        self._log("\n你可以检查左侧匹配列表无误后，再点击【开始合并】。")

    def _merge(self) -> None:
        v = self._validate_dirs()
        if not v:
            return

        try:
            pmin, pmax = self._get_range()
        except ValueError as e:
            messagebox.showerror("范围输入错误", str(e))
            return

        desc_dir, trans_dir, out_dir = v

        if not self._last_scan:
            self._scan()
            if not self._last_scan:
                return

        desc_scan, trans_scan = self._last_scan

        # Recompute common with range (avoid stale)
        desc_ids = set(filter_indices(sorted(desc_scan.mapping.keys()), pmin, pmax))
        trans_ids = set(filter_indices(sorted(trans_scan.mapping.keys()), pmin, pmax))
        common = sorted(desc_ids & trans_ids)

        if not common:
            messagebox.showwarning("无法合并", "在当前 P 范围内没有找到可匹配的【P序号】文件对。请检查范围或目录内容。")
            return

        overwrite = self.overwrite.get()
        add_headers = self.add_headers.get()
        order = self.order.get()

        self._log("\n==============================")
        self._log("开始合并写入输出文件...")
        self._log(f"P范围：{('全部' if (pmin is None and pmax is None) else f'P{pmin or ''} ~ P{pmax or ''}')}")

        out_dir.mkdir(parents=True, exist_ok=True)

        merged_count = 0
        failed_count = 0

        for idx in common:
            desc_path = choose_most_complete(desc_scan.mapping[idx])
            trans_path = choose_most_complete(trans_scan.mapping[idx])

            chosen_name = trans_path.name if len(trans_path.stem) >= len(desc_path.stem) else desc_path.name
            out_path = safe_output_path(out_dir, chosen_name, overwrite=overwrite)

            try:
                merged = merge_pair(
                    desc_path=desc_path,
                    transcript_path=trans_path,
                    order=order,
                    add_headers=add_headers,
                )
                out_path.write_text(merged, encoding="utf-8-sig")  # Windows Notepad friendly
                merged_count += 1
            except Exception as e:
                failed_count += 1
                self._log(f"[失败] P{idx} 合并失败：{e}")

        # Also log unmatched (range-limited) for export traceability
        only_desc = sorted(desc_ids - trans_ids)
        only_trans = sorted(trans_ids - desc_ids)

        self._log(f"\n合并完成 ✅")
        self._log(f"- 成功输出：{merged_count} 个文件")
        self._log(f"- 合并失败：{failed_count} 个文件")
        self._log(f"- 输出目录：{out_dir}")

        if only_desc or only_trans:
            self._log("\n[未匹配清单（当前范围）]")
            if only_desc:
                self._log(f"仅简介存在：{only_desc}")
            if only_trans:
                self._log(f"仅转写存在：{only_trans}")

        messagebox.showinfo("完成", f"合并完成！\n成功输出：{merged_count}\n失败：{failed_count}\n输出目录：\n{out_dir}")

    def _open_out_dir(self) -> None:
        v = self._validate_dirs()
        if not v:
            return
        _, _, out_dir = v
        try:
            open_folder(out_dir)
        except Exception as e:
            messagebox.showerror("打开目录失败", str(e))

    def _export_log(self) -> None:
        text = self.log.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("无内容", "日志为空，无需导出。")
            return

        # Suggest default filename
        default_name = "merge_log.txt"
        path = filedialog.asksaveasfilename(
            title="导出日志为 TXT",
            defaultextension=".txt",
            initialfile=default_name,
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        if not path:
            return

        try:
            Path(path).write_text(text + "\n", encoding="utf-8-sig")
            messagebox.showinfo("导出成功", f"日志已导出：\n{path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))


def main() -> None:
    # Windows: improve DPI scaling
    try:
        if sys.platform.startswith("win"):
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
