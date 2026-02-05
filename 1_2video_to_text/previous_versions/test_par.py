import os
import time
from tkinter import Tk, filedialog, Button, Label, Listbox, Scrollbar, END, SINGLE, messagebox
from openai import OpenAI
from dotenv import load_dotenv
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

# 加载环境变量
load_dotenv()
client = OpenAI()

MAX_THREADS = 3

# 单个视频转换函数（含计时）
def transcribe_video(video_path):
    file_start = time.time()
    file_name = os.path.basename(video_path)
    transcript_file_name = os.path.splitext(file_name)[0] + ".txt"

    try:
        with open(video_path, "rb") as audio_file:
            transcript = client.audio.transcriptions.create(
                model="gpt-4o-mini-transcribe",
                file=audio_file,
            )

        with open(transcript_file_name, "w", encoding='utf-8') as f:
            f.write(transcript.text)

        duration = time.time() - file_start
        return file_name, f"Success in {duration:.2f}s"

    except Exception as e:
        return file_name, f"Error: {e}"

# 线程处理函数（含总计时）
def transcribe_videos(file_paths):
    start_time = time.time()
    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        future_to_video = {executor.submit(transcribe_video, path): path for path in file_paths}

        for future in as_completed(future_to_video):
            file_name, result = future.result()
            update_status(f"{file_name}: {result}")

    total_duration = time.time() - start_time
    update_status(f"All tasks completed in {total_duration:.2f}s")
    transcribe_btn.config(state='normal')

# 文件选择回调
def select_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Video Files", "*.mp4 *.mov *.avi")])
    if file_paths:
        listbox.delete(0, END)
        for path in file_paths:
            listbox.insert(END, path)
        transcribe_btn.config(state='normal')

# 转换按钮回调（启动线程）
def start_transcription():
    files = listbox.get(0, END)
    if files:
        transcribe_btn.config(state='disabled')
        threading.Thread(target=transcribe_videos, args=(files,), daemon=True).start()
    else:
        messagebox.showwarning("No Files", "Please select files to transcribe.")

# 状态更新函数
def update_status(message):
    status_listbox.insert(END, message)
    status_listbox.yview(END)

# GUI设置
root = Tk()
root.title("Concurrent Transcription with Timing")

Label(root, text="Selected Videos:").pack(pady=5)

scrollbar = Scrollbar(root)
scrollbar.pack(side='right', fill='y')

listbox = Listbox(root, selectmode=SINGLE, width=80, yscrollcommand=scrollbar.set)
listbox.pack(pady=5)
scrollbar.config(command=listbox.yview)

select_btn = Button(root, text="Select Video Files", command=select_files)
select_btn.pack(pady=10)

transcribe_btn = Button(root, text="Transcribe", command=start_transcription, state='disabled')
transcribe_btn.pack(pady=10)

Label(root, text="Processing Status:").pack(pady=5)
status_listbox = Listbox(root, width=80)
status_listbox.pack(pady=5)

root.mainloop()
