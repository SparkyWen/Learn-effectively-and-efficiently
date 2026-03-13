import os
import time
import threading
from tkinter import (
    Tk, filedialog, Button, Label, Listbox, Scrollbar,
    END, SINGLE, messagebox, Entry, IntVar
)
from openai import OpenAI
from dotenv import load_dotenv, set_key, dotenv_values
from concurrent.futures import ThreadPoolExecutor, as_completed

# =========================
# 基础配置
# =========================
dotenv_path = ".env"
output_folder = "transcription_results"
os.makedirs(output_folder, exist_ok=True)

# 支持的媒体文件类型
SUPPORTED_MEDIA_PATTERNS = [
    "*.mp4", "*.mov", "*.avi",
    "*.mp3", "*.m4a", "*.wav", "*.aac", "*.flac",
    "*.ogg", "*.mpga", "*.mpeg", "*.webm"
]

# 全局 client
client = None

# 加载 .env
load_dotenv(dotenv_path)


# =========================
# OpenAI Client
# =========================
def get_client():
    api_key = dotenv_values(dotenv_path).get("OPENAI_API_KEY")
    if api_key:
        return OpenAI(api_key=api_key)
    else:
        messagebox.showerror("Error", "API Key not set!")
        return None


# =========================
# UI 安全更新函数
# =========================
def update_status(message):
    root.after(0, _append_status, message)


def _append_status(message):
    status_listbox.insert(END, message)
    status_listbox.yview(END)


def set_buttons_state(transcribing: bool):
    def _update():
        if transcribing:
            transcribe_btn.config(state='disabled')
            select_btn.config(state='disabled')
        else:
            transcribe_btn.config(state='normal')
            select_btn.config(state='normal')
    root.after(0, _update)


def show_warning(title, msg):
    root.after(0, lambda: messagebox.showwarning(title, msg))


def show_error(title, msg):
    root.after(0, lambda: messagebox.showerror(title, msg))


def show_info(title, msg):
    root.after(0, lambda: messagebox.showinfo(title, msg))


# =========================
# 单文件转写
# =========================
def transcribe_media(media_path):
    global client

    thread_name = threading.current_thread().name
    file_name = os.path.basename(media_path)
    transcript_file_name = os.path.join(
        output_folder,
        os.path.splitext(file_name)[0] + ".txt"
    )

    print(f"[{thread_name}] Starting transcription of {file_name} at {time.strftime('%H:%M:%S')}")
    file_start = time.time()

    try:
        with open(media_path, "rb") as media_file:
            transcript = client.audio.transcriptions.create(
                model="gpt-4o-mini-transcribe",
                file=media_file,
                chunking_strategy="auto",
            )

        with open(transcript_file_name, "w", encoding="utf-8") as f:
            f.write(transcript.text)

        duration = time.time() - file_start
        print(f"[{thread_name}] Finished transcription of {file_name} in {duration:.2f}s at {time.strftime('%H:%M:%S')}")
        return file_name, f"Success in {duration:.2f}s (Thread: {thread_name})"

    except Exception as e:
        print(f"[{thread_name}] Error transcribing {file_name}: {e}")
        return file_name, f"Error: {e} (Thread: {thread_name})"


# =========================
# 批量转写线程函数
# =========================
def transcribe_files(file_paths, max_workers):
    global client

    start_time = time.time()

    try:
        with ThreadPoolExecutor(
            max_workers=max_workers,
            thread_name_prefix='Worker'
        ) as executor:
            future_to_file = {
                executor.submit(transcribe_media, path): path for path in file_paths
            }

            for future in as_completed(future_to_file):
                file_name, result = future.result()
                update_status(f"{file_name}: {result}")

        total_duration = time.time() - start_time
        update_status(f"All tasks completed in {total_duration:.2f}s")
        show_info("Done", f"All files completed in {total_duration:.2f}s")

    except Exception as e:
        show_error("Processing Error", str(e))
        update_status(f"Batch processing failed: {e}")

    finally:
        set_buttons_state(False)


# =========================
# 文件选择
# =========================
def select_files():
    file_paths = filedialog.askopenfilenames(
        title="Select Media Files",
        filetypes=[
            ("Media Files", " ".join(SUPPORTED_MEDIA_PATTERNS)),
            ("Video Files", "*.mp4 *.mov *.avi"),
            ("Audio Files", "*.mp3 *.m4a *.wav *.aac *.flac *.ogg *.mpga *.mpeg *.webm"),
            ("All Files", "*.*"),
        ]
    )

    if file_paths:
        listbox.delete(0, END)
        for path in file_paths:
            listbox.insert(END, path)

        status_listbox.insert(END, f"Selected {len(file_paths)} file(s).")
        status_listbox.yview(END)
        transcribe_btn.config(state='normal')


# =========================
# 启动转写
# =========================
def start_transcription():
    global client

    files = listbox.get(0, END)
    if not files:
        messagebox.showwarning("No Files", "Please select media files to transcribe.")
        return

    try:
        max_workers = int(thread_entry.get().strip())
        if max_workers <= 0:
            raise ValueError
    except ValueError:
        messagebox.showwarning("Invalid Threads", "Please enter a valid positive integer for Max Threads.")
        return

    client = get_client()
    if not client:
        return

    status_listbox.delete(0, END)
    update_status(f"Starting transcription for {len(files)} file(s) with {max_workers} thread(s)...")

    set_buttons_state(True)
    threading.Thread(
        target=transcribe_files,
        args=(list(files), max_workers),
        daemon=True
    ).start()


# =========================
# 保存 API Key
# =========================
def save_api_key():
    api_key = api_entry.get().strip()
    if api_key:
        set_key(dotenv_path, "OPENAI_API_KEY", api_key)
        messagebox.showinfo("Saved", "API Key saved successfully!")
    else:
        messagebox.showwarning("Empty", "API Key cannot be empty.")


# =========================
# GUI
# =========================
root = Tk()
root.title("Concurrent Media Transcription with Custom API & Thread Count")
root.geometry("760x700")

Label(root, text="Enter OpenAI API Key:").pack(pady=5)
api_entry = Entry(root, width=80)
api_entry.pack(pady=5)

# 如果 .env 里已有 key，则自动回填
existing_api_key = dotenv_values(dotenv_path).get("OPENAI_API_KEY", "")
if existing_api_key:
    api_entry.insert(0, existing_api_key)

Button(root, text="Save API Key", command=save_api_key).pack(pady=5)

Label(root, text="Set Max Threads:").pack(pady=5)
thread_count = IntVar(value=10)
thread_entry = Entry(root, textvariable=thread_count, width=10)
thread_entry.pack(pady=5)

Label(root, text="Selected Media Files:").pack(pady=5)

scrollbar = Scrollbar(root)
scrollbar.pack(side='right', fill='y')

listbox = Listbox(root, selectmode=SINGLE, width=100, height=12, yscrollcommand=scrollbar.set)
listbox.pack(pady=5)
scrollbar.config(command=listbox.yview)

select_btn = Button(root, text="Select Media Files", command=select_files)
select_btn.pack(pady=10)

transcribe_btn = Button(root, text="Transcribe", command=start_transcription, state='disabled')
transcribe_btn.pack(pady=10)

Label(root, text="Processing Status:").pack(pady=5)
status_listbox = Listbox(root, width=100, height=15)
status_listbox.pack(pady=5)

root.mainloop()