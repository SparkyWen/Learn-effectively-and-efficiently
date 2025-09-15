
# Excel 合并助手（.xlsx）

这个小工具可以把一个文件夹中的所有 `.xlsx` 文件合并到一个大的 `.xlsx` 文件里。支持 GUI、并行读取加速、两种合并模式，以及一些常用的数据清洗与写入选项。

## 功能特点
- **合并为一个总表**：把所有文件中的所有 Sheet 纵向堆叠到一个工作表（超过 1,048,576 行自动分页写到 `merged_1/merged_2/...`）。
- **按 Sheet 名分别合并**：对同名 Sheet 各自进行合并，输出多个工作表。
- **并行读取**：线程池 / 进程池两种方式，文件多时显著加速。
- **实用选项**：递归子目录、添加来源文件/Sheet 列、去空行、列并集/交集、按照指定列去重、自动列宽。
- **健壮性**：自动跳过 Excel 临时文件（`~$xxx.xlsx`），列名重复自动重命名（`_1/_2`）。

## 安装与运行
1. 安装 Python 3.9+（Windows 自带 tkinter，一般无需额外安装）。
2. 在命令行进入本目录，执行：
   ```bash
   python -m venv venv
   venv\Scripts\activate  # Windows
   # 或 source venv/bin/activate （macOS/Linux）
   pip install -r requirements.txt
   python excel_merger_gui.py
   ```
   Windows 用户也可以直接双击运行 `run_windows.bat`。

## 使用建议
- **写入引擎**：默认 `openpyxl`。若安装了 `XlsxWriter`，可切换为 `xlsxwriter`，写入速度通常更快。
- **列集合策略**：数据结构不完全一致时，选择“并集”可保留所有列（缺失填 NaN），选择“交集”只保留所有表均存在的列。
- **Excel 行上限**：若合并后超过 1,048,576 行，工具会自动分页写入多个工作表。
- **极大数据**：如数据非常大且 Excel 难以承载，建议改为输出 CSV 或数据库，以获得更好的可扩展性。

## 常见问题
- **GUI 无响应？** 合并任务在后台线程执行，正常会保持界面响应；如数据量过大，耐心等待即可。
- **进程池在 Windows 下报错？** 请确保直接运行 `python excel_merger_gui.py`（而不是交互式解释器里 `import`），并使用 Python 3.9+。
- **中文路径/文件名**：脚本使用 UTF-8 并与 tkinter GUI 配合，一般可正确处理中文路径。
