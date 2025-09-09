import os
from tkinter import Tk, filedialog, Button, Label, Listbox, Scrollbar, END, SINGLE, messagebox
from openai import OpenAI
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()
client = OpenAI()

# 创建转换函数
def transcribe_videos(file_paths):
    for video_path in file_paths:
        try:
            file_name = os.path.basename(video_path)
            transcript_file_name = os.path.splitext(file_name)[0] + ".txt"
            
            with open(video_path, "rb") as audio_file:
                transcript = client.audio.transcriptions.create(
                    model="gpt-4o-mini-transcribe",
                    file=audio_file,
                )

            # 保存转录文件
            with open(transcript_file_name, "w", encoding='utf-8') as f:
                f.write(transcript.text)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while processing {file_name}: {e}")

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
        transcribe_videos(files)
        messagebox.showinfo("Success", "Transcription completed successfully!")
    else:
        messagebox.showwarning("No Files", "Please select files to transcribe.")

# GUI设置
root = Tk()
root.title("Batch Video to Text Transcription")

Label(root, text="Selected Videos:").pack(pady=5)

frame = Scrollbar(root)
frame.pack(side='right', fill='y')

listbox = Listbox(root, selectmode=SINGLE, width=80, yscrollcommand=frame.set)
listbox.pack(pady=5)
frame.config(command=listbox.yview)

select_btn = Button(root, text="Select Video Files", command=select_files)
select_btn.pack(pady=10)

transcribe_btn = Button(root, text="Transcribe", command=start_transcription, state='disabled')
transcribe_btn.pack(pady=10)

root.mainloop()
