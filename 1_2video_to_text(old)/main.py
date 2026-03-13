import os
import time
import threading
from tkinter import (
    Tk, filedialog, Button, Label, Listbox, Scrollbar,
    END, SINGLE, messagebox, Entry, StringVar, IntVar
)
from openai import OpenAI
from dotenv import load_dotenv, set_key, dotenv_values
from concurrent.futures import ThreadPoolExecutor, as_completed

# 环境文件路径
dotenv_path = ".env"

# 确保结果文件夹存在
output_folder = "transcription_results"
os.makedirs(output_folder, exist_ok=True)

# 加载API key
load_dotenv(dotenv_path)
def get_client():
    api_key = dotenv_values(dotenv_path).get("OPENAI_API_KEY")
    if api_key:
        return OpenAI(api_key=api_key)
    else:
        messagebox.showerror("Error", "API Key not set!")
        return None

# 单个视频转换函数（含详细计时和线程信息）
def transcribe_video(video_path):
    thread_name = threading.current_thread().name
    file_name = os.path.basename(video_path)
    transcript_file_name = os.path.join(output_folder, os.path.splitext(file_name)[0] + ".txt")

    print(f"[{thread_name}] Starting transcription of {file_name} at {time.strftime('%H:%M:%S')}")

    file_start = time.time()
    try:
        with open(video_path, "rb") as audio_file:
            transcript = client.audio.transcriptions.create(
                model="gpt-4o-mini-transcribe",
                file=audio_file,
            )

        with open(transcript_file_name, "w", encoding='utf-8') as f:
            f.write(transcript.text)

        duration = time.time() - file_start

        print(f"[{thread_name}] Finished transcription of {file_name} in {duration:.2f}s at {time.strftime('%H:%M:%S')}")
        return file_name, f"Success in {duration:.2f}s (Thread: {thread_name})"

    except Exception as e:
        print(f"[{thread_name}] Error transcribing {file_name}: {e}")
        return file_name, f"Error: {e} (Thread: {thread_name})"

# 线程处理函数（含总计时）
def transcribe_videos(file_paths):
    global client
    client = get_client()
    if not client:
        transcribe_btn.config(state='normal')
        return

    start_time = time.time()
    with ThreadPoolExecutor(max_workers=thread_count.get(), thread_name_prefix='Worker') as executor:
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

# 保存API key回调
def save_api_key():
    api_key = api_entry.get().strip()
    if api_key:
        set_key(dotenv_path, "OPENAI_API_KEY", api_key)
        messagebox.showinfo("Saved", "API Key saved successfully!")
    else:
        messagebox.showwarning("Empty", "API Key cannot be empty.")

# GUI设置
root = Tk()
root.title("Concurrent Transcription with Custom API & Thread Count")

Label(root, text="Enter OpenAI API Key:").pack(pady=5)
api_entry = Entry(root, width=80)
api_entry.pack(pady=5)
Button(root, text="Save API Key", command=save_api_key).pack(pady=5)

Label(root, text="Set Max Threads:").pack(pady=5)
thread_count = IntVar(value=10)
thread_entry = Entry(root, textvariable=thread_count, width=10)
thread_entry.pack(pady=5)

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