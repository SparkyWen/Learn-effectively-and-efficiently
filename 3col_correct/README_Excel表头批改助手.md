
# Excel 表头批改助手（.xlsx）

本工具用于**批量修改 Excel 文件（.xlsx）首行表头**，其它内容（数据、公式、样式）尽量保持不变。

## 功能
- 图形界面（Tkinter），易上手
- 并行加速：线程/进程两种方式
- 表头自定义（逗号/换行分隔）
- 选择处理 **所有 Sheet / 仅第一个 / 指定名称**
- 输出方式：原地覆盖（可选 .bak 备份）或输出到新目录（默认安全）
- 三种写入策略：
  1. **只覆盖已有列（不改变列数）**（推荐）
  2. 覆盖已有列 + 多余列标题清空
  3. **强制按目标列写满**（可能超出现有列数）

## 安装 & 运行
```bash
python -m venv venv
# Windows: venv\Scripts\activate
# macOS/Linux: source venv/bin/activate
pip install -r requirements.txt
python excel_header_renamer_gui.py
```
Windows 用户可直接双击 `run_header_renamer_windows.bat` 自动创建虚拟环境并启动。

## 常见建议
- **推荐输出到新目录**，避免意外覆盖。原地覆盖时可勾选“生成备份”。
- 若文件很多/很大，优先尝试**进程模式**并发。
- 表头目标列与实际列数不一致时，优先使用“只覆盖已有列”。
