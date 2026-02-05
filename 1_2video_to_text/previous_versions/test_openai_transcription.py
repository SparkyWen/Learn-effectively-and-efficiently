import os
from openai import OpenAI
from dotenv import load_dotenv
load_dotenv()
client = OpenAI()

audio_file = open("./SP/【P101】AI大模型周报2025年5月d（DeepSeek-R1-0528Claude4Gemma.mp4", "rb")
transcript = client.audio.transcriptions.create(
  model="gpt-4o-mini-transcribe",
  file=audio_file,
)
with open("transcript.txt", "w") as f:
    f.write(transcript.text)