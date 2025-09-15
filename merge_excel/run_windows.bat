
@echo off
setlocal
title Excel 合并助手 - 启动器
if not exist venv (
    echo [*] 正在创建虚拟环境 venv ...
    python -m venv venv
)
call venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python excel_merger_gui.py
