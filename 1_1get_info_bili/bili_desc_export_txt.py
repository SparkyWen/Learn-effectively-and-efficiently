#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Bilibili 合集：批量导出“每个视频自己的简介”到独立 txt 文件（带 GUI）

用法：
  python bili_desc_export_gui.py

说明：
- 在输入框里粘贴：合集内任意一集的 URL 或 BV 号
- 选择输出目录，点击【开始导出】
- 每个视频生成一个 txt，即使没有简介也会生成空文件
- 文件命名： 【P1】标题.txt、【P2】标题.txt ...
- 如果遇到风控/需要登录权限，可在界面里粘贴 Cookie（例如包含 SESSDATA）
"""

from __future__ import annotations

import os
import re
import sys
import time
import threading
import queue
from typing import Dict, List, Optional, Tuple

import requests
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText


UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/122.0.0.0 Safari/537.36"
)


# -----------------------------
# 基础工具
# -----------------------------

def extract_bvid(url_or_bvid: str) -> str:
    s = (url_or_bvid or "").strip()
    if s.startswith("BV"):
        return s
    m = re.search(r"(BV[a-zA-Z0-9]+)", s)
    if not m:
        raise ValueError("输入中未找到 BV 号。请粘贴包含 BV 的视频链接或直接输入 BVxxxxxxxx。")
    return m.group(1)


def safe_filename(name: str, max_len: int = 180) -> str:
    """
    清理 Windows 文件名非法字符，并限制长度。
    Windows 禁止：<>:"/\\|?* 以及控制字符；末尾不能是空格/点
    """
    name = name or ""
    name = re.sub(r"[\x00-\x1f\x7f]", "", name)  # 控制字符

    name = name.replace("\\", "＼").replace("/", "／")
    name = re.sub(r'[<>:"|?*]', " ", name)

    name = re.sub(r"\s+", " ", name).strip()
    name = name.rstrip(" .")

    if len(name) > max_len:
        name = name[:max_len].rstrip(" .")

    return name if name else "无标题"


def ensure_unique_path(path: str) -> str:
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    k = 1
    while True:
        new_path = f"{base} ({k}){ext}"
        if not os.path.exists(new_path):
            return new_path
        k += 1


# -----------------------------
# B站接口逻辑（方案A）
# -----------------------------

def build_session(cookie: Optional[str] = None) -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "User-Agent": UA,
        "Referer": "https://www.bilibili.com/",
        "Origin": "https://www.bilibili.com",
    })
    if cookie:
        s.headers.update({"Cookie": cookie})
    return s


def get_view(session: requests.Session, bvid: str) -> Dict:
    url = "https://api.bilibili.com/x/web-interface/view"
    r = session.get(url, params={"bvid": bvid}, timeout=20)
    r.raise_for_status()
    j = r.json()
    if j.get("code") != 0:
        raise RuntimeError(f"view 接口失败: bvid={bvid} code={j.get('code')} msg={j.get('message')}")
    return j["data"]


def get_desc_text(view_data: Dict) -> str:
    """
    更稳的简介提取：
    - 优先 desc（常见）
    - 若 desc 为空，再尝试 desc_v2（分段）
    """
    desc = view_data.get("desc")
    if isinstance(desc, str) and desc.strip():
        return desc

    dv2 = view_data.get("desc_v2") or []
    parts: List[str] = []
    if isinstance(dv2, list):
        for seg in dv2:
            if not isinstance(seg, dict):
                continue
            txt = seg.get("raw_text") or seg.get("text") or ""
            if isinstance(txt, str) and txt:
                parts.append(txt)

    return "\n".join(parts).strip()


def get_collection_meta(session: requests.Session, any_bvid: str) -> Tuple[int, int]:
    data = get_view(session, any_bvid)
    ugc = data.get("ugc_season")
    if not ugc or not isinstance(ugc, dict) or "id" not in ugc:
        raise RuntimeError("该视频返回中没有 ugc_season：可能不是“合集”形态，或需要登录/权限。")
    mid = int(data["owner"]["mid"])
    season_id = int(ugc["id"])
    return mid, season_id


def list_collection_items(session: requests.Session, mid: int, season_id: int, page_size: int = 30) -> List[Dict]:
    endpoints = [
        "https://api.bilibili.com/x/polymer/web-space/seasons_archives_list",
        "https://api.bilibili.com/x/polymer/space/seasons_archives_list",  # 兜底
    ]

    all_items: List[Dict] = []
    page_num = 1
    total: Optional[int] = None

    while True:
        data = None
        last_err = None

        for ep in endpoints:
            try:
                r = session.get(ep, params={
                    "mid": mid,
                    "season_id": season_id,
                    "page_num": page_num,
                    "page_size": page_size,
                    "sort_reverse": "false",
                }, timeout=20)
                r.raise_for_status()
                j = r.json()
                if j.get("code") != 0:
                    raise RuntimeError(f"code={j.get('code')} msg={j.get('message')}")
                data = j.get("data") or {}
                break
            except Exception as e:
                last_err = e

        if data is None:
            raise RuntimeError(f"合集列表接口失败：{last_err}")

        archives = data.get("archives") or data.get("items") or []
        if not isinstance(archives, list):
            archives = []

        all_items.extend([x for x in archives if isinstance(x, dict)])

        page = data.get("page") or {}
        if total is None:
            t = page.get("total")
            if isinstance(t, int):
                total = t

        if total is not None and len(all_items) >= total:
            break
        if len(archives) < page_size:
            break

        page_num += 1
        time.sleep(0.25)

    return all_items


# -----------------------------
# 导出逻辑（写 txt）
# -----------------------------

def export_descriptions_to_txt(
    input_url_or_bvid: str,
    out_dir: str,
    sleep_sec: float,
    cookie: str,
    stop_event: threading.Event,
    ui_queue: "queue.Queue[tuple]"
) -> None:
    """
    通过 ui_queue 向 UI 线程发送事件：
      ("log", "xxx")
      ("progress_init", total)
      ("progress", current, total)
      ("done", out_dir)
      ("error", "errmsg")
      ("stopped",)
    """
    def log(msg: str):
        ui_queue.put(("log", msg))

    try:
        os.makedirs(out_dir, exist_ok=True)
        bvid = extract_bvid(input_url_or_bvid)

        session = build_session(cookie=cookie.strip() or None)

        log(f"输入 BV：{bvid}")
        mid, season_id = get_collection_meta(session, bvid)
        log(f"解析到 mid={mid}, season_id={season_id}")

        items = list_collection_items(session, mid, season_id, page_size=30)

        ordered: List[Dict] = []
        seen = set()
        for it in items:
            bv = it.get("bvid")
            if isinstance(bv, str) and bv.startswith("BV") and bv not in seen:
                seen.add(bv)
                ordered.append(it)

        total = len(ordered)
        ui_queue.put(("progress_init", total))
        log(f"合集视频数：{total}")

        for idx, it in enumerate(ordered, 1):
            if stop_event.is_set():
                ui_queue.put(("stopped",))
                return

            bvid_i = it.get("bvid") or ""
            fallback_title = it.get("title") or bvid_i or "无标题"

            title = fallback_title
            desc_text = ""

            try:
                view_data = get_view(session, bvid_i)
                t = view_data.get("title")
                if isinstance(t, str) and t.strip():
                    title = t.strip()
                desc_text = get_desc_text(view_data)
            except Exception as e:
                log(f"[{idx}/{total}] {bvid_i} 获取失败：{e} -> 仍生成空简介文件")

            safe_title = safe_filename(title)
            filename = f"【P{idx}】{safe_title}.txt"
            path = os.path.join(out_dir, filename)

            # 防止路径过长（Windows 常见）
            if len(path) > 240:
                safe_title = safe_filename(title, max_len=80)
                filename = f"【P{idx}】{safe_title}.txt"
                path = os.path.join(out_dir, filename)

            path = ensure_unique_path(path)

            with open(path, "w", encoding="utf-8", newline="\n") as f:
                f.write(desc_text or "")

            preview = (desc_text[:60].replace("\n", " ") if desc_text else "")
            log(f"[{idx}/{total}] {bvid_i} desc_len={len(desc_text)} -> {os.path.basename(path)} | {preview!r}")

            ui_queue.put(("progress", idx, total))
            time.sleep(max(0.0, sleep_sec))

        ui_queue.put(("done", out_dir))

    except Exception as e:
        ui_queue.put(("error", str(e)))


# -----------------------------
# GUI
# -----------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("B站合集简介批量导出（每集单独TXT）")
        self.geometry("900x650")

        self.ui_queue: "queue.Queue[tuple]" = queue.Queue()
        self.worker_thread: Optional[threading.Thread] = None
        self.stop_event = threading.Event()

        self._build_ui()
        self.after(80, self._poll_queue)

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill="x", **pad)

        # 输入 URL / BV
        ttk.Label(frm, text="合集任意一集 URL 或 BV：").grid(row=0, column=0, sticky="w")
        self.var_input = tk.StringVar()
        ent = ttk.Entry(frm, textvariable=self.var_input)
        ent.grid(row=0, column=1, sticky="we", padx=(8, 0))
        frm.columnconfigure(1, weight=1)

        # 输出目录
        ttk.Label(frm, text="输出目录：").grid(row=1, column=0, sticky="w")
        self.var_out = tk.StringVar(value=os.path.abspath("bili_desc_txt"))
        ent_out = ttk.Entry(frm, textvariable=self.var_out)
        ent_out.grid(row=1, column=1, sticky="we", padx=(8, 0))

        btn_browse = ttk.Button(frm, text="选择...", command=self._choose_dir)
        btn_browse.grid(row=1, column=2, padx=(8, 0))

        # 请求间隔
        ttk.Label(frm, text="请求间隔(秒)：").grid(row=2, column=0, sticky="w")
        self.var_sleep = tk.StringVar(value="0.2")
        sp = ttk.Spinbox(frm, from_=0.0, to=10.0, increment=0.1, textvariable=self.var_sleep, width=10)
        sp.grid(row=2, column=1, sticky="w", padx=(8, 0))

        # Cookie（可选）
        cookie_box = ttk.LabelFrame(self, text="可选：Cookie（风控/权限需要登录时才填）")
        cookie_box.pack(fill="both", expand=False, padx=10, pady=8)

        self.txt_cookie = ScrolledText(cookie_box, height=4)
        self.txt_cookie.pack(fill="x", padx=10, pady=8)
        self.txt_cookie.insert("1.0", "")  # 默认空

        # 控制按钮
        ctrl = ttk.Frame(self)
        ctrl.pack(fill="x", padx=10, pady=6)

        self.btn_start = ttk.Button(ctrl, text="开始导出", command=self._start)
        self.btn_start.pack(side="left")

        self.btn_stop = ttk.Button(ctrl, text="停止", command=self._stop, state="disabled")
        self.btn_stop.pack(side="left", padx=(10, 0))

        self.btn_clear = ttk.Button(ctrl, text="清空日志", command=self._clear_log)
        self.btn_clear.pack(side="left", padx=(10, 0))

        # 进度条
        prog = ttk.Frame(self)
        prog.pack(fill="x", padx=10, pady=6)

        self.var_prog = tk.IntVar(value=0)
        self.var_total = tk.IntVar(value=0)

        self.progress = ttk.Progressbar(prog, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", expand=True, side="left")

        self.lbl_prog = ttk.Label(prog, text="0/0")
        self.lbl_prog.pack(side="left", padx=(10, 0))

        # 日志区
        log_box = ttk.LabelFrame(self, text="运行日志")
        log_box.pack(fill="both", expand=True, padx=10, pady=8)

        self.txt_log = ScrolledText(log_box, height=18)
        self.txt_log.pack(fill="both", expand=True, padx=10, pady=8)
        self._log("准备就绪。粘贴合集任意一集的URL/BV，选择输出目录，点击“开始导出”。")

    def _log(self, msg: str):
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")

    def _clear_log(self):
        self.txt_log.delete("1.0", "end")

    def _choose_dir(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.var_out.set(d)

    def _start(self):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("提示", "任务正在运行中。")
            return

        input_val = self.var_input.get().strip()
        if not input_val:
            messagebox.showerror("错误", "请输入合集任意一集的 URL 或 BV。")
            return

        out_dir = self.var_out.get().strip()
        if not out_dir:
            messagebox.showerror("错误", "请选择输出目录。")
            return

        try:
            sleep_sec = float(self.var_sleep.get().strip())
            if sleep_sec < 0:
                raise ValueError
        except Exception:
            messagebox.showerror("错误", "请求间隔必须是非负数字，例如 0.2")
            return

        cookie = self.txt_cookie.get("1.0", "end").strip()

        self.stop_event.clear()
        self._set_running(True)

        self._log("=" * 70)
        self._log("开始执行...")
        self._log(f"输出目录：{os.path.abspath(out_dir)}")
        self._log(f"请求间隔：{sleep_sec}s")
        self._log(f"Cookie：{'已填写' if cookie else '未填写'}")

        self.worker_thread = threading.Thread(
            target=export_descriptions_to_txt,
            args=(input_val, out_dir, sleep_sec, cookie, self.stop_event, self.ui_queue),
            daemon=True
        )
        self.worker_thread.start()

    def _stop(self):
        if self.worker_thread and self.worker_thread.is_alive():
            self.stop_event.set()
            self._log("收到停止请求：将在当前视频处理完后安全停止...")

    def _set_running(self, running: bool):
        self.btn_start.configure(state="disabled" if running else "normal")
        self.btn_stop.configure(state="normal" if running else "disabled")

    def _poll_queue(self):
        try:
            while True:
                item = self.ui_queue.get_nowait()

                if not item:
                    continue

                typ = item[0]

                if typ == "log":
                    self._log(item[1])

                elif typ == "progress_init":
                    total = int(item[1])
                    self.var_total.set(total)
                    self.var_prog.set(0)
                    self.progress["maximum"] = max(1, total)
                    self.progress["value"] = 0
                    self.lbl_prog.configure(text=f"0/{total}")

                elif typ == "progress":
                    cur, total = int(item[1]), int(item[2])
                    self.var_prog.set(cur)
                    self.progress["value"] = cur
                    self.lbl_prog.configure(text=f"{cur}/{total}")

                elif typ == "done":
                    out_dir = item[1]
                    self._log("=" * 70)
                    self._log(f"完成 ✅ 输出目录：{os.path.abspath(out_dir)}")
                    self._set_running(False)
                    messagebox.showinfo("完成", f"导出完成！\n输出目录：\n{os.path.abspath(out_dir)}")

                elif typ == "error":
                    self._log("=" * 70)
                    self._log(f"失败 ❌ {item[1]}")
                    self._set_running(False)
                    messagebox.showerror("失败", item[1])

                elif typ == "stopped":
                    self._log("=" * 70)
                    self._log("已停止（用户中断）。")
                    self._set_running(False)
                    messagebox.showinfo("已停止", "任务已停止。")

        except queue.Empty:
            pass

        self.after(80, self._poll_queue)


def main():
    # Windows 上一些环境会对 stdout 编码敏感，这里尽量不改动
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()