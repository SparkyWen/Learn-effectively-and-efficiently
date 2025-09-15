import os
import time
from tkinter import Tk, filedialog, Button, Label, Listbox, Scrollbar, END, SINGLE, messagebox
from openai import OpenAI
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()
client = OpenAI()

# 创建转换函数
def transcribe_videos(file_paths):
    start_time = time.time()  # 开始计时
    for video_path in file_paths:
        file_start = time.time()
        try:
            file_name = os.path.basename(video_path)
            transcript_file_name = os.path.splitext(file_name)[0] + ".txt"
            
            with open(video_path, "rb") as audio_file:
                transcript = client.audio.transcriptions.create(
                    model="gpt-4o-mini-transcribe",
                    file=audio_file,
                )

            with open(transcript_file_name, "w", encoding='utf-8') as f:
                f.write(transcript.text)

            duration = time.time() - file_start
            status_listbox.insert(END, f"{file_name} completed in {duration:.2f}s")
            status_listbox.yview(END)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred processing {file_name}: {e}")

    total_duration = time.time() - start_time
    status_listbox.insert(END, f"All files completed in {total_duration:.2f}s")
    status_listbox.yview(END)

# 文件选择回调
def select_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Video Files", "*.mp4 *.mov *.avi")])
    if file_paths:
        listbox.delete(0, END)
        for path in file_paths:
            listbox.insert(END, path)
        transcribe_btn.config(state='normal')

# 转换按钮回调
def start_transcription():
    files = listbox.get(0, END)
    if files:
        transcribe_btn.config(state='disabled')
        transcribe_videos(files)
        transcribe_btn.config(state='normal')
    else:
        messagebox.showwarning("No Files", "Please select files to transcribe.")

# GUI设置
root = Tk()
root.title("Sequential Transcription with Timing")

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
