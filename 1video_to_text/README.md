# 并发视频转写工具（Tkinter GUI + OpenAI Transcription）

本项目提供一个**带图形界面（GUI）**的批量转写工具：你可以一次选择多个视频文件（`.mp4/.mov/.avi`），设置并发线程数，然后调用 OpenAI 的转写接口将音频内容转成文字，并把结果保存为 `.txt` 文件。

---

## 目录

- [功能概览](#功能概览)
- [脚本说明](#脚本说明)
- [输入与输出](#输入与输出)
- [运行环境与依赖](#运行环境与依赖)
- [安装与配置](#安装与配置)
- [使用方法](#使用方法)
- [并发与性能说明](#并发与性能说明)
- [常见问题（FAQ）](#常见问题-faq)
- [安全建议](#安全建议)

---

## 功能概览

- ✅ **图形界面操作**：无需命令行参数，点选文件即可运行  
- ✅ **批量转写**：一次选择多个视频文件
- ✅ **并发处理**：使用 `ThreadPoolExecutor` 多线程并发提交转写请求
- ✅ **线程数可配置**：GUI 可设置最大线程数（默认 10）
- ✅ **状态可视化**：GUI 列表中实时追加每个文件的结果（成功/失败/耗时/线程信息）
- ✅ **结果自动落盘**：每个视频输出一个同名 `.txt` 转写文件到 `transcription_results/`

---

## 脚本说明

> 你提供的脚本是一个 **单文件 GUI 应用**，核心逻辑可以理解为以下模块：

### 1) 环境与输出目录初始化
- 读取/写入环境文件：`.env`
- 自动创建输出目录：`transcription_results/`

```python
dotenv_path = ".env"
output_folder = "transcription_results"
os.makedirs(output_folder, exist_ok=True)
```

### 2) OpenAI Client 初始化（读取 API Key）

- 通过 `python-dotenv` 从 `.env` 中读取 `OPENAI_API_KEY`
- 若未配置，会弹窗报错并停止任务

```
def get_client():
    api_key = dotenv_values(dotenv_path).get("OPENAI_API_KEY")
    ...
    return OpenAI(api_key=api_key)
```

### 3) 单文件转写函数 `transcribe_video(video_path)`

对每个视频文件执行以下步骤：

1. 打开文件（二进制方式）
2. 调用 OpenAI 转写接口（模型：`gpt-4o-mini-transcribe`）
3. 将返回的 `transcript.text` 写入本地 `.txt`
4. 捕获异常并记录错误信息

输出文件名规则：

- 输入：`xxx.mp4`
- 输出：`transcription_results/xxx.txt`

### 4) 多文件并发转写 `transcribe_videos(file_paths)`

- 从 GUI 获取线程数：`thread_count.get()`
- `ThreadPoolExecutor(max_workers=...)` 并发提交多个 `transcribe_video`
- `as_completed` 按完成顺序收集结果并更新 GUI 状态栏
- 统计总耗时，完成后恢复按钮可点击状态

### 5) GUI 交互逻辑

- **Select Video Files**：弹出文件选择框，多选视频，写入“Selected Videos”列表
- **Transcribe**：启动后台线程执行 `transcribe_videos`，避免 GUI 卡死
- **Save API Key**：把用户输入的 Key 写入 `.env` 中的 `OPENAI_API_KEY`

------

## 输入与输出

### 输入（你需要提供什么）

1. **OpenAI API Key**
   - 在 GUI 的输入框中填写并点击 **Save API Key**
   - 或手动创建 `.env` 文件（见下方配置说明）
2. **视频文件**
   - GUI 中点击 **Select Video Files**
   - 支持筛选扩展名：`*.mp4 *.mov *.avi`

> 注意：转写能力本质依赖视频文件中**音轨**是否可读，以及接口对封装/编码的支持情况。遇到报错时可先将视频转为 `.mp3/.wav` 再试（见 FAQ）。

### 输出（会得到什么）

- 默认输出目录：`./transcription_results/`
- 每个视频生成一个对应 `.txt` 文件：
  - `a.mp4` → `transcription_results/a.txt`
  - `lecture.mov` → `transcription_results/lecture.txt`

此外：

- GUI 的“Processing Status”会持续追加每个文件的执行结果
- 控制台（终端）会打印更详细的线程名与时间戳日志

------

## 运行环境与依赖

### Python 版本

建议：**Python 3.9+**（更高版本也可）

### 依赖库

- `openai`（新版 SDK，支持 `from openai import OpenAI` 用法）
- `python-dotenv`
- `tkinter`（多数系统自带；Windows/macOS 通常默认可用，部分 Linux 需要额外安装）

------

## 安装与配置

### 1) 安装依赖

```
pip install openai python-dotenv
```

> 如果你有多个 Python 环境（conda/venv），强烈建议用：

```
python -m pip install openai python-dotenv
```

确保安装到**当前运行脚本的同一个解释器环境**里。

### 2) 配置 API Key（两种方式）

#### 方式 A：GUI 内保存（推荐）

1. 运行脚本打开 GUI
2. 在 “Enter OpenAI API Key” 输入框粘贴你的 Key
3. 点击 **Save API Key**
4. 程序会自动在当前目录创建/更新 `.env`

#### 方式 B：手动创建 `.env`

在脚本同目录新建 `.env` 文件，内容如下：

```
OPENAI_API_KEY=你的key写在这里
```

> `.env` 必须与脚本在**同一工作目录**下（脚本里写死为 `dotenv_path = ".env"`）。

------

## 使用方法

### 运行脚本

假设脚本文件名为 `transcribe_gui.py`：

```
python transcribe_gui.py
```

### GUI 操作步骤

1. **填写并保存 API Key**
   - 粘贴 Key → 点击 **Save API Key**
2. **设置最大线程数**
   - “Set Max Threads” 默认是 10
   - 如果你担心触发速率限制/网络不稳，建议改成 2~5
3. **选择视频文件**
   - 点击 **Select Video Files** → 多选视频 → 确认
4. **开始转写**
   - 点击 **Transcribe**
5. **查看结果**
   - GUI 状态栏会显示每个文件是否成功、耗时、线程信息
   - 转写文本在 `transcription_results/` 下

------

## 并发与性能说明

这个脚本使用的是**多线程并发**策略：

- `ThreadPoolExecutor(max_workers=N)` 同时处理 N 个视频
- 对每个视频都会发起一次转写请求（I/O 密集型任务，多线程可以提升吞吐）
- 线程数越大并不一定越快，原因包括：
  - API 侧可能存在速率限制（Rate Limit）
  - 网络带宽与本机磁盘 I/O
  - 视频越大上传越慢

**经验建议：**

- 小文件、网络稳定：`5~10`
- 文件很大/网络一般：`2~5`
- 经常报错（429/超时）：降线程数 + 分批处理

------

## 常见问题（FAQ）

### 1) 明明装了 openai，但运行提示 `No module named openai`

这通常是因为你安装依赖的 Python 环境与运行脚本的环境不是同一个。建议：

- 用 `python -m pip install openai python-dotenv`
- 在运行脚本前确认 `python` 指向的解释器就是你装库的那个

### 2) 提示 “API Key not set!”

说明 `.env` 中没有 `OPENAI_API_KEY`，或脚本运行时找不到 `.env`。
 排查：

- `.env` 是否在**当前运行目录**（不是脚本所在目录就一定能找到，取决于你从哪里启动）
- `.env` 是否包含：
  - `OPENAI_API_KEY=...`

### 3) 转写报错：文件格式/编码不支持、读取失败

可能原因：视频封装/音频编码不兼容，或文件过大/网络不稳。
 建议：

- 先用 ffmpeg 把视频提取为音频再试（示例）：

```
ffmpeg -i input.mp4 -vn -acodec mp3 output.mp3
```

然后你可以把文件选择框的筛选扩展名改成支持 `.mp3/.wav`（需要改代码）。

### 4) 频繁出现 429 / rate limit / 超时

建议：

- 降低线程数（比如 10 → 3）
- 分批选择文件（一次 10 个、20 个）
- 确保网络稳定

### 5) 为什么 GUI 没卡死？

因为点击 **Transcribe** 后，脚本会创建一个后台线程执行 `transcribe_videos`：

```
threading.Thread(target=transcribe_videos, args=(files,), daemon=True).start()
```

这样 Tkinter 主线程还能继续响应界面刷新与操作。

------

## 安全建议

- **不要把 `.env` 上传到 GitHub**
   建议在 `.gitignore` 中加入：

  ```
  .env
  transcription_results/
  ```

- API Key 属于敏感信息，尽量不要截图分享或发给他人。

- 使用 OpenAI API 会产生费用，请自行关注用量与账单。