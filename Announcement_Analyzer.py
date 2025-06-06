#Announcement_Analyzer - アナ朗評価

import tkinter as tk
from tkinter import filedialog, messagebox
from pydub import AudioSegment
import speech_recognition as sr
import difflib
import librosa
import numpy as np
import os
import time
import pykakasi
import re
from tkinter import ttk
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tempfile
import simpleaudio as sa
import sys
import ctypes
ctypes.windll.shcore.SetProcessDpiAwareness(1)

kks = pykakasi.kakasi()
latest_file_path = None

root = tk.Tk()
root.withdraw() 
icon_path = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), 'icon.ico')
root.iconbitmap(icon_path)

root.title("Announcement_Analyzer - アナ朗評価")
root.geometry("1050x950")
font_mono = ("Courier New", 11)  

mode = tk.StringVar(value="アナウンス")

pitch_window = None

mode_frame = tk.Frame(root)
mode_frame.pack(fill="x", padx=10, pady=(10, 0))

mode_left_frame = tk.Frame(mode_frame)
mode_left_frame.pack(side="left")

tk.Label(mode_left_frame, text="  評価モード：").pack(side="left")
tk.Radiobutton(mode_left_frame, text="アナウンス", variable=mode, value="アナウンス").pack(side="left")
tk.Radiobutton(mode_left_frame, text="朗読", variable=mode, value="朗読").pack(side="left")


def clean_transcription(text):
    endings = ["ました", "ます", "でした", "です"]
    for ending in endings:
        pattern = rf'({ending}) (?=[^。])'
        replacement = r'\1。\n'
        text = re.sub(pattern, replacement, text)
    text = text.replace(" ", "").replace("　", "")
    return text

def update_progress(percent, message):
    progress_label.config(text=f"{percent}%：{message}")
    root.update_idletasks()

def clear_results():
    labels = [
        volume_val, volume_score, volume_comment,
        match_val, match_score, match_comment,
        speed_val, speed_score, speed_comment,
        pitch_val, pitch_score, pitch_comment
    ]
    for label in labels:
        label.config(text="")
    transcription_box.delete("1.0", tk.END)

def count_reading_syllables(text):
    text = text.replace("\n", "")  
    result = kks.convert(text)
    reading = ''.join([item['hira'] for item in result])
    return len(reading)

def select_file():
    global latest_file_path
    file_path = filedialog.askopenfilename(filetypes=[("WAV files", "*.wav")])
    if file_path:
        latest_file_path = file_path
        messagebox.showinfo("選択完了", f"選択されたファイル: {os.path.basename(file_path)}")
        try:
            update_progress(20, "音声ファイルを読み込み中...")
            duration = librosa.get_duration(path=file_path)
            mp3_status_label.config(text=f"WAV：{os.path.basename(file_path)}（{duration:.2f}秒）")

            update_progress(40, "声の大きさを解析中...")
            analyze_volume(file_path)

            update_progress(75, "滑舌と読みスピードを解析中...")
            recognize_audio(file_path, duration)

            update_progress(100, "解析完了")
            table_frame.pack(pady=10)
        except Exception as e:
            messagebox.showerror("エラー", f"音声ファイルの処理に失敗しました:\n{e}")

def transcribe_audio():
    global latest_file_path
    if not latest_file_path:
        messagebox.showwarning("警告", "先にWAVファイルを選択してください")
        return
    try:
        recognizer = sr.Recognizer()
        with sr.AudioFile(latest_file_path) as source:
            audio_data = recognizer.record(source)
            recognized_text = recognizer.recognize_google(audio_data, language="ja-JP")
            cleaned = clean_transcription(recognized_text)
            transcription_box.delete("1.0", tk.END)
            transcription_box.insert(tk.END, cleaned)
    except Exception as e:
        messagebox.showerror("エラー", f"文字起こしに失敗しました:\n{e}")

def compare_transcription_with_original():
    transcription_box.tag_remove("mismatch", "1.0", tk.END)
    original = text_input.get("1.0", tk.END).strip()
    recognized = transcription_box.get("1.0", tk.END).strip()
    matcher = difflib.SequenceMatcher(None, original, recognized)
    opcodes = matcher.get_opcodes()
    for tag, i1, i2, j1, j2 in opcodes:
        if tag != 'equal':
            start = f"1.0 + {j1} chars"
            end = f"1.0 + {j2} chars"
            transcription_box.tag_add("mismatch", start, end)

def analyze_volume(wav_path):
    audio = AudioSegment.from_wav(wav_path)
    samples = np.array(audio.get_array_of_samples()).astype(np.float32)

    if audio.channels == 2:
        samples = samples.reshape((-1, 2)).mean(axis=1)

    rms = np.sqrt(np.mean(samples ** 2))

    ref = 1.0
    db = 20 * np.log10(rms / ref + 1e-10)
    db = max(0, db)  

    score = max(0, min(100, (db / 80) * 100))

    if score >= 90:
        comment = "とても大きな声"
    elif score >= 80:
        comment = "大きな声"
    elif score >= 70:
        comment = "まあまあ大きい"
    elif score >= 60:
        comment = "ちょっと小さめ"
    else:
        comment = "かなり小さめ"

    volume_val.config(text=f"：平均 {db:.2f} dB")
    volume_score.config(text=f"スコア {score:>4.0f}/100")
    volume_comment.config(text="　　" + comment)

def recognize_audio(wav_path, duration_sec):
    recognizer = sr.Recognizer()
    with sr.AudioFile(wav_path) as source:
        audio_data = recognizer.record(source)
        try:
            recognized_text = recognizer.recognize_google(audio_data, language="ja-JP")
            original_text = text_input.get("1.0", tk.END).strip()
            matcher = difflib.SequenceMatcher(None, original_text, recognized_text)
            match_percentage = matcher.ratio() * 100

            if match_percentage >= 90:
                comment = "すごくはっきり"
            elif match_percentage >= 85:
                comment = "かなりはっきり"
            elif match_percentage >= 80:
                comment = "まあまあ分かる"
            elif match_percentage >= 75:
                comment = "なんとか分かる"
            elif match_percentage >= 70:
                comment = "聞き取りにくい"
            else:
                comment = "聞き取れない"

            match_val.config(text=f"：{match_percentage:.2f}％")
            match_score.config(text=f"スコア {match_percentage:>4.0f}/100")
            match_comment.config(text="　　" + comment)

            syllables = count_reading_syllables(recognized_text)
            spm = syllables / (duration_sec / 60)

            score, comment = get_speed_score_comment(spm, mode.get())

            speed_val.config(text=f"：{spm:.2f} 音/分")
            speed_score.config(text=f"スコア {score:>4.0f}/200")
            speed_comment.config(text="　　" + comment)

        except sr.UnknownValueError:
            match_val.config(text="：---")
            match_score.config(text="スコア ----")
            match_comment.config(text="　　認識できません")
            speed_val.config(text="：---")
            speed_score.config(text="スコア ----")
            speed_comment.config(text="　　測定不能")
        except sr.RequestError as e:
            match_val.config(text="：APIエラー")
            speed_val.config(text="：APIエラー")

def get_speed_score_comment(spm, mode):
    stats = {
        "アナウンス": {"mean": 333.7851, "std": 30.84654},
        "朗読": {"mean": 259.4353, "std": 30.3969}
    }

    mean = stats[mode]["mean"]
    std = stats[mode]["std"]

    deviation = spm - mean
    z = deviation / std  

    score = 100 + z * 30
    score = round(max(0, min(200, score)), 2)

    if score < 60:
        comment = "けっこう遅め"
    elif score < 85:
        comment = "ちょっと遅め"
    elif score <= 115:
        comment = "理想的な速さ"
    elif score <= 140:
        comment = "ちょっと速め"
    else:
        comment = "けっこう速い"

    return score, comment

def save_text():
    file_path = filedialog.asksaveasfilename(defaultextension=".txt")
    if file_path:
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(text_input.get("1.0", tk.END))

def load_text():
    file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    if file_path:
        with open(file_path, "r", encoding="utf-8") as f:
         content = f.read().rstrip()  
         text_input.delete("1.0", tk.END)
         text_input.insert(tk.END, content)

def analyze_text():
    text = text_input.get("1.0", tk.END).strip()
    word_count = len(text.split())

    text = text.replace("\n", "").replace("。", "")

    hira_text = ''.join([item['hira'] for item in kks.convert(text)])
    char_count = len(hira_text)

    text_status_label.config(text=f"文字数（ひらがな換算）：{char_count}　文章数：{word_count}")

tk.Label(root, text="原稿を入力：").pack(pady=(10, 0))
text_input_frame = tk.Frame(root)
text_input_frame.pack()
text_scroll = tk.Scrollbar(text_input_frame)
text_scroll.pack(side="right", fill="y")
text_input = tk.Text(text_input_frame, height=10, width=90, yscrollcommand=text_scroll.set, font=font_mono)
text_input.pack()
text_scroll.config(command=text_input.yview)

button_frame = tk.Frame(root)
button_frame.pack(pady=5)
tk.Button(button_frame, text=" 原稿を解析 ", command=analyze_text).pack(side="left", padx=5)
tk.Button(button_frame, text=" 原稿を読み込み ", command=load_text).pack(side="left", padx=5)
tk.Button(button_frame, text=" 原稿を保存 ", command=save_text).pack(side="left", padx=5)

text_status_label = tk.Label(root, text="文字数：---　語数：---")
text_status_label.pack()

tk.Label(root, text="文字起こし（赤字：原稿と不一致）").pack(pady=(10, 0))
transcription_frame = tk.Frame(root)
transcription_frame.pack()
transcription_scroll = tk.Scrollbar(transcription_frame)
transcription_scroll.pack(side="right", fill="y")
transcription_box = tk.Text(transcription_frame, height=10, width=90, yscrollcommand=transcription_scroll.set, font=font_mono)
transcription_box.pack()
transcription_box.tag_configure("mismatch", foreground="#FF0055")
transcription_scroll.config(command=transcription_box.yview)

transcription_button_frame = tk.Frame(root)
transcription_button_frame.pack(pady=(5, 10))
tk.Button(transcription_button_frame, text=" 文字起こし ", command=transcribe_audio).pack(side="left", padx=5)
tk.Button(transcription_button_frame, text=" 文字起こしと原稿を比較 ", command=compare_transcription_with_original).pack(side="left", padx=5)

progress_label = tk.Label(root, text="進行状況：未開始")
progress_label.pack(pady=(10, 0))
tk.Button(root, text=" WAVファイルを選択して評価 ", command=select_file).pack(pady=10)
mp3_status_label = tk.Label(root, text="WAVファイル：未選択")
mp3_status_label.pack()

table_frame = tk.Frame(root)

def add_row(row, title, val_label, score_label, comment_label):
    tk.Label(table_frame, text=title, font=font_mono, width=15, anchor="e").grid(row=row, column=0, sticky="e")
    val_label.config(font=font_mono, width=28, anchor="w")
    val_label.grid(row=row, column=1, sticky="w")
    score_label.config(font=font_mono, width=14, anchor="w")
    score_label.grid(row=row, column=2, sticky="w")
    comment_label.config(font=font_mono, fg="gray", anchor="w")
    comment_label.grid(row=row, column=3, sticky="w")

volume_val, volume_score, volume_comment = tk.Label(table_frame), tk.Label(table_frame), tk.Label(table_frame)
match_val, match_score, match_comment = tk.Label(table_frame), tk.Label(table_frame), tk.Label(table_frame)
speed_val, speed_score, speed_comment = tk.Label(table_frame), tk.Label(table_frame), tk.Label(table_frame)
pitch_val, pitch_score, pitch_comment = tk.Label(table_frame), tk.Label(table_frame), tk.Label(table_frame)

add_row(0, "声の大きさ", volume_val, volume_score, volume_comment)
add_row(1, "滑舌（一致率）", match_val, match_score, match_comment)
add_row(2, "読みのスピード", speed_val, speed_score, speed_comment)
  

root.deiconify() 
root.mainloop()





#Announcement_Analyzer - アナ朗ピッチ分析

import tkinter as tk 
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
from pydub import AudioSegment
import librosa
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tempfile
import os
import simpleaudio as sa
import time
import pykakasi
import sys

class PitchViewer:
    def __init__(self, root):
        self.root = root
        self.root.withdraw() 
        icon_path = os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), 'icon.ico')
        self.root.iconbitmap(icon_path)
        self.root.title("Announcement_Analyzer - アナ朗ピッチ分析")  

        self.default_font = ("TkDefaultFont", 11)
        self.root.option_add("*Font", self.default_font)

        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        self.main_frame.columnconfigure(0, weight=10)
        self.main_frame.columnconfigure(1, weight=0)
        self.main_frame.rowconfigure(2, weight=3)
        self.main_frame.rowconfigure(3, weight=0)    

        left_top_frame = ttk.Frame(self.main_frame)
        left_top_frame.grid(row=0, column=0, sticky="nw", padx=10, pady=10)    

        self.button = ttk.Button(left_top_frame, text="WAVを選択して解析", command=self.load_and_analyze)
        self.button.pack(anchor="w")    

        self.script_button = ttk.Button(left_top_frame, text="原稿を読み込み", command=self.load_script)
        self.script_button.pack(anchor="w", pady=(5, 0))    

        self.info_label = ttk.Label(left_top_frame, text="ファイル名　-秒　文章数：-")
        self.info_label.pack(anchor="w", pady=(5, 0))    

        self.script_sentences = []    

        bottom_top_frame = ttk.Frame(self.main_frame)
        bottom_top_frame.grid(row=1, column=0, sticky="w", padx=10, pady=(0, 5))    

        self.sentence_label = ttk.Label(bottom_top_frame, text="表示文番号：")
        self.sentence_label.pack(side=tk.LEFT)    

        self.sentence_var = tk.StringVar()
        self.sentence_combo = ttk.Combobox(bottom_top_frame, textvariable=self.sentence_var, width=5, state="readonly")
        self.sentence_combo.pack(side=tk.LEFT, padx=(5, 0))    

        self.go_button = ttk.Button(bottom_top_frame, text="表示", command=self.display_selected_sentence)
        self.go_button.pack(side=tk.LEFT, padx=(5, 0))    

        self.play_button = ttk.Button(bottom_top_frame, text="再生", command=self.play_current_sentence)
        self.play_button.pack(side=tk.LEFT, padx=(10, 0))    

        self.duration_label = ttk.Label(bottom_top_frame, text="表示中の文章：-秒")
        self.duration_label.pack(side=tk.LEFT, padx=(10, 0))

        self.plot_frame = ttk.Frame(self.main_frame)
        self.plot_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)    

        self.script_frame = ttk.Frame(self.main_frame)
        self.script_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 10))
        self.script_label = ttk.Label(self.script_frame, text="　", anchor="center", wraplength=1000, justify="center", relief="ridge")
        self.script_label.pack(fill=tk.X)    

        right_side_frame = ttk.Frame(self.main_frame, width=280)
        right_side_frame.grid(row=0, column=1, rowspan=4, sticky="nsew", padx=(0, 10), pady=10)
        right_side_frame.grid_propagate(False)
        right_side_frame.columnconfigure(0, weight=1)    

        mode_frame = ttk.LabelFrame(right_side_frame, text="文の区切りの判定方法")
        mode_frame.grid(row=0, column=0, sticky="new", ipadx=5, ipady=5)    

        self.mode_var = tk.StringVar(value="auto")
        self.auto_radio = ttk.Radiobutton(mode_frame, text="自動（差分最大点）", variable=self.mode_var, value="auto", command=self.update_threshold_state)
        self.auto_radio.grid(row=0, column=0, sticky="w")
        self.auto_threshold_var = tk.StringVar(value="-")
        self.auto_threshold_entry = ttk.Entry(mode_frame, textvariable=self.auto_threshold_var, width=6, state='readonly')
        self.auto_threshold_entry.grid(row=0, column=1, padx=(10, 0))    

        self.manual_radio = ttk.Radiobutton(mode_frame, text="手動（しきい値）", variable=self.mode_var, value="manual", command=self.update_threshold_state)
        self.manual_radio.grid(row=1, column=0, sticky="w")
        self.manual_threshold_var = tk.DoubleVar(value=1.5)
        self.manual_threshold_entry = ttk.Entry(mode_frame, textvariable=self.manual_threshold_var, width=6)
        self.manual_threshold_entry.grid(row=1, column=1, padx=(10, 0))    

        self.fixed_radio = ttk.Radiobutton(mode_frame, text="原稿あり（文数指定）", variable=self.mode_var, value="fixed", command=self.update_threshold_state)
        self.fixed_radio.grid(row=2, column=0, sticky="w")
        self.fixed_sentence_count_var = tk.IntVar(value=6)
        self.fixed_sentence_entry = ttk.Entry(mode_frame, textvariable=self.fixed_sentence_count_var, width=6)
        self.fixed_sentence_entry.grid(row=2, column=1, padx=(10, 0))    

        self.reanalyze_button = ttk.Button(mode_frame, text="再解析", command=self.reanalyze)
        self.reanalyze_button.grid(row=3, column=0, columnspan=2, pady=(5, 0))        

        self.pause_frame = ttk.LabelFrame(right_side_frame, text="無音区間リスト")
        self.pause_frame.grid(row=2, column=0, sticky="nsew")
        self.pause_frame.grid_remove()
        self.pause_frame.columnconfigure(0, weight=1)
        self.pause_frame.rowconfigure(0, weight=1)    

        frame_sort_toggle = ttk.Frame(right_side_frame)
        frame_sort_toggle.grid(row=1, column=0, sticky="w", pady=(10, 0))        

        self.show_pauses = tk.BooleanVar(value=False)
        self.toggle_pause_button = ttk.Checkbutton(frame_sort_toggle, text="無音区間リストを表示", variable=self.show_pauses, command=self.toggle_pause_frame)
        self.toggle_pause_button.pack(side=tk.LEFT)        

        self.sort_mode = tk.StringVar(value="番号順")
        self.sort_combo = ttk.Combobox(
            frame_sort_toggle,
            textvariable=self.sort_mode,
            values=["番号順", "長い順"],
            state="readonly",
            width=8
        )
        self.sort_combo.pack(side=tk.LEFT, padx=(10, 0))
        self.sort_combo.bind("<<ComboboxSelected>>", lambda e: self.update_pause_text())    

        self.auto_pitch_range = tk.BooleanVar(value=True)
        self.pitch_range_checkbox = ttk.Checkbutton(
            bottom_top_frame,
            text="ピッチの範囲を文章ごとに自動調整",
            variable=self.auto_pitch_range,
            command=self.display_selected_sentence
        )
        self.pitch_range_checkbox.pack(side=tk.LEFT, padx=(10, 0))   

        self.data_by_sentence = []
        self.raw_audio = None
        self.sr = None
        self.intervals = []
        self.display_empty_graph()

        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(100, lambda: self.root.attributes("-topmost", False))

        self.kakasi = pykakasi.kakasi()
        self.converter = self.kakasi

    def update_threshold_state(self):
        self.auto_threshold_entry.config(state='readonly')
        self.manual_threshold_entry.config(state='disabled')
        self.fixed_sentence_entry.config(state='disabled')
        if self.mode_var.get() == "manual":
            self.manual_threshold_entry.config(state='normal')
        elif self.mode_var.get() == "fixed":
            self.fixed_sentence_entry.config(state='normal')
    
    def load_and_analyze(self):
        file_path = filedialog.askopenfilename(filetypes=[("WAV files", "*.wav")], parent=self.root)
        if not file_path:
            return
    
        messagebox.showinfo("選択完了", f"選択されたファイル: {os.path.basename(file_path)}", parent=self.root)
    
        filename = os.path.basename(file_path)

        duration = librosa.get_duration(path=file_path)
        y, sr = librosa.load(file_path, sr=None)
    
        self.raw_audio = y
        self.sr = sr
    
        self.pause_text = tk.Text(self.pause_frame, wrap='none', font=("Courier New", 11))
        self.pause_text.grid(row=0, column=0, sticky="nsew")
        self.pause_text.config(state='disabled')
    
        self.info_label.config(text=f"{filename}　{duration:.2f}秒")
        self.reanalyze()

    def reanalyze(self):
        if self.raw_audio is None or self.sr is None:
            return
    
        y = self.raw_audio
        sr = self.sr

        intervals = librosa.effects.split(y, top_db=30)
        self.all_intervals = intervals  
        pauses = []
        for i in range(len(intervals) - 1):
            prev_end = intervals[i][1]
            next_start = intervals[i + 1][0]
            pause_length = (next_start - prev_end) / sr
            if pause_length > 0.10:
                pauses.append((i, pause_length))
    
        boundary_indices = []
        threshold = 1.5
    
        if self.mode_var.get() == "auto":
            sorted_pauses = sorted(pauses, key=lambda x: x[1], reverse=True)
            pause_lengths = [p[1] for p in sorted_pauses]
            max_drop = 0
            boundary_index = 0
            for i in range(len(pause_lengths) - 1):
                diff = pause_lengths[i] - pause_lengths[i + 1]
                if diff > max_drop:
                    max_drop = diff
                    boundary_index = i + 1
            threshold = pause_lengths[boundary_index]
            self.auto_threshold_var.set(f"{threshold:.2f}")
          
            boundary_indices = sorted([p[0] for p in sorted_pauses[:boundary_index]])
    
        elif self.mode_var.get() == "manual":
            threshold = self.manual_threshold_var.get()
            self.auto_threshold_var.set("-")
            boundary_indices = [i for i, p in pauses if p >= threshold]
    
        elif self.mode_var.get() == "fixed":
            sentence_count = self.fixed_sentence_count_var.get()
            top_n = max(1, sentence_count - 1)
            sorted_pauses = sorted(pauses, key=lambda x: x[1], reverse=True)
            boundary_indices = [i for i, _ in sorted_pauses[:top_n]]
            boundary_indices.sort()
            self.auto_threshold_var.set("-")

        self.boundary_indices = boundary_indices  
    
        sentence_intervals = []
        start_idx = 0
        for idx in boundary_indices:
            sentence_intervals.append((intervals[start_idx][0], intervals[idx][1]))
            start_idx = idx + 1
        if start_idx < len(intervals):
            sentence_intervals.append((intervals[start_idx][0], intervals[-1][1]))
    
        self.intervals = sentence_intervals
        self.data_by_sentence.clear()
    
        for (start, end) in sentence_intervals:
            y_segment = y[start:end]
            t_start = start / sr
    
            f0, _, _ = librosa.pyin(
                y_segment,
                fmin=librosa.note_to_hz('C2'),
                fmax=librosa.note_to_hz('C7'),
                sr=sr,
                frame_length=2048,
                hop_length=512
            )
    
            if f0 is not None and np.sum(~np.isnan(f0)) > 0:
                times = librosa.times_like(f0, sr=sr, hop_length=512) + t_start
                f0 = np.where((f0 >= 100) & (f0 <= 400), f0, np.nan)
                rms = librosa.feature.rms(y=y_segment, frame_length=2048, hop_length=512)[0]
    
                self.data_by_sentence.append({
                    "start_time": t_start,
                    "times": times,
                    "pitches": f0,
                    "rms": rms
                })

        all_pitches = []
        for data in self.data_by_sentence:
            if data["pitches"] is not None:
                all_pitches.extend(data["pitches"][~np.isnan(data["pitches"])])
        if all_pitches:
            self.global_pitch_min = max(0, np.min(all_pitches) - 20)
            self.global_pitch_max = np.max(all_pitches) + 20
        else:
            self.global_pitch_min = 100
            self.global_pitch_max = 400
    
        self.pauses = pauses
        self.update_pause_text()
    
        self.info_label.config(text=self.info_label.cget("text").split("　文章数：")[0] + f"　文章数：{len(self.data_by_sentence)}")
        self.display_selected_sentence()
    
        self.sentence_combo['values'] = [str(i + 1) for i in range(len(self.data_by_sentence))]
        self.sentence_combo.current(0)

    def toggle_pause_frame(self):
        if self.show_pauses.get():
            self.pause_frame.grid()
        else:
            self.pause_frame.grid_remove()

    def display_selected_sentence(self):
        try:
            idx = int(self.sentence_combo.get()) - 1
        except ValueError:
            return
    
        if idx < 0 or idx >= len(self.data_by_sentence):
            return
    
        for widget in self.plot_frame.winfo_children():
            widget.destroy()
    
        data = self.data_by_sentence[idx]
        times = data["times"]
        f0 = data["pitches"]
        rms = data["rms"]
    
        sentence_start = data["start_time"]
        sentence_end = data["times"][-1]
    
        threshold_orange = np.percentile(rms, 80)
        threshold_pink = np.percentile(rms, 70)
    
        fig, ax = plt.subplots(figsize=(10, 4))
    
        self.span_artists = []

        extra_band = None
        if idx < len(self.intervals) - 1:
            this_end = self.intervals[idx][1]
            next_start = self.intervals[idx + 1][0]
            for pause_index, (i, pause_len) in enumerate(self.pauses):
                if i + 1 >= len(self.all_intervals):
                    continue
                pause_start = self.all_intervals[i][1]
                pause_end = self.all_intervals[i + 1][0]
                if this_end <= pause_start <= next_start:
                    extra_band = (pause_start / self.sr, pause_end / self.sr)
                    self.span_artists.append((None, pause_index, extra_band[0], extra_band[1]))  
                    break
    
        if extra_band:
            xlim_right = extra_band[1]
        else:
            xlim_right = times[-1]
        ax.set_xlim(times[0], xlim_right)

        if self.auto_pitch_range.get():
            valid_pitches = f0[~np.isnan(f0)]
            if len(valid_pitches) > 0:
                min_pitch = valid_pitches.min()
                max_pitch = valid_pitches.max()
                margin = max(max_pitch - min_pitch, 50) * 0.2
                pitch_ylim = (max(0, min_pitch - margin), max_pitch + margin)
            else:
                pitch_ylim = (100, 400)
        else:
            pitch_ylim = (self.global_pitch_min, self.global_pitch_max)
        
        ax.set_ylim(*pitch_ylim)
    
        ax.set_ylabel("Pitch (Hz)")
        ax.set_xlabel("Time (s)")
        ax.grid(True)
    
        for global_index, (i, pause_length) in enumerate(self.pauses):
            if i + 1 >= len(self.all_intervals):
                continue
            pause_start = self.all_intervals[i][1] / self.sr
            pause_end = self.all_intervals[i + 1][0] / self.sr
            if pause_end >= sentence_start and pause_start <= sentence_end:
                span = ax.axvspan(pause_start, pause_end, color='#2EEAD1', alpha=0.3)
                self.span_artists.append((span, global_index, pause_start, pause_end))

        for i in range(len(self.span_artists)):
            span, index, start, end = self.span_artists[i]
            if span is None:
                new_span = ax.axvspan(start, end, color='#2EEAD1', alpha=0.3)
                self.span_artists[i] = (new_span, index, start, end)
    
        def is_in_silent_band(t):
            for _, _, start, end in self.span_artists:
                if start <= t <= end:
                    return True
            return False
    
        for i in range(len(times) - 1):
            if np.isnan(f0[i]) or np.isnan(f0[i + 1]):
                continue
            t0, t1 = times[i], times[i + 1]
            if is_in_silent_band((t0 + t1) / 2):
                continue
            if rms[i] >= threshold_orange:
                color = '#FF0055'
            elif rms[i] >= threshold_pink:
                color = '#990033'
            else:
                color = '#000000'
            ax.plot([t0, t1], [f0[i], f0[i + 1]], color=color)
    
        self.playback_line = ax.axvline(times[0], color='#FF0044', linewidth=1)
        self.canvas = FigureCanvasTkAgg(fig, master=self.plot_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        plt.close(fig)

        duration_sec = sentence_end - sentence_start
        
        if 0 <= idx < len(self.script_sentences):
            sentence = self.script_sentences[idx]
            hira = ''.join([item['hira'] for item in self.converter.convert(sentence)]) 
            hira = hira.replace(" ", "").replace("\n", "")
            char_count = len(hira)
        else:
            char_count = 0

        if 0 <= idx < len(self.script_sentences) and duration_sec > 0 and char_count > 0:
            speed = char_count / duration_sec
            self.duration_label.config(text=f"表示中の文章：{duration_sec:.2f}秒　速さ：{speed:.2f}文字/秒")
        else:
            self.duration_label.config(text=f"表示中の文章：{duration_sec:.2f}秒")

        if 0 <= idx < len(self.script_sentences):
            self.script_label.config(text=self.script_sentences[idx])
        else:
            self.script_label.config(text="")
    
        self.canvas.mpl_connect('button_press_event', self.on_graph_click)
       
    def play_current_sentence(self):
        try:
            idx = int(self.sentence_combo.get()) - 1
        except ValueError:
            return    

        if idx < 0 or idx >= len(self.data_by_sentence):
            return    

        start_sample = int(self.intervals[idx][0])
        end_sample = int(self.intervals[idx][1])    

        y_segment = self.raw_audio[start_sample:end_sample]
        y_segment = (y_segment * 32767).astype(np.int16)    

        self.play_obj = sa.play_buffer(y_segment, 1, 2, self.sr)
        self.start_time = time.time()
        self.current_sentence_index = idx  
        self.monitor_playback()

    def monitor_playback(self):
        if not hasattr(self, 'playback_line') or not self.play_obj.is_playing():
            return    

        elapsed = time.time() - self.start_time
        absolute_time = self.data_by_sentence[self.current_sentence_index]["start_time"] + elapsed
        self.playback_line.set_xdata([absolute_time, absolute_time])
        self.canvas.draw()
        self.root.after(50, self.monitor_playback)

    def load_script(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")],parent=self.root)
        
        if not file_path:
            return   

        messagebox.showinfo("読み込み完了", f"選択された原稿ファイル: {os.path.basename(file_path)}",parent=self.root) 
    
        with open(file_path, "r", encoding="utf-8") as f:
            content = f.read()    
    
        raw_sentences = content.replace("。", "。\n").splitlines()
        sentences = [line.strip() for line in raw_sentences if line.strip()]    
    
        self.script_sentences = sentences
        self.fixed_sentence_count_var.set(len(sentences))
        self.mode_var.set("fixed")
        self.update_threshold_state()
    
        print(f"[原稿読み込み完了] {len(sentences)}文")
    
        self.reanalyze() 


    def display_empty_graph(self):
        for widget in self.plot_frame.winfo_children():
            widget.destroy()
        fig, ax = plt.subplots(figsize=(10, 4))
        ax.set_xlim(0, 10)
        ax.set_ylim(100, 400)
        ax.set_ylabel("Pitch (Hz)")
        ax.set_xlabel("Time (s)")
        ax.grid(True)
        canvas = FigureCanvasTkAgg(fig, master=self.plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        plt.close(fig)

    def update_pause_text(self):
        if not hasattr(self, 'pauses'):
            return
    
        sort_by_length = self.sort_mode.get() == "長い順"
        self.pause_text.config(state='normal')
        self.pause_text.delete('1.0', tk.END)
    
        max_number_len = len(str(len(self.pauses)))
    
        if sort_by_length:
            self.sorted_pauses_for_display = sorted(enumerate(self.pauses), key=lambda x: x[1][1], reverse=True)
            sentence_count = len(self.intervals)
            divider_index = sentence_count - 1
        else:
            self.sorted_pauses_for_display = list(enumerate(self.pauses))
            divider_index = None
    
        self.pause_index_to_line_index = {}  
        
        line_index = 0
        for i, (original_index, pause) in enumerate(self.sorted_pauses_for_display):
            number_str = f"{original_index + 1:>{max_number_len}}"
            self.pause_text.insert(tk.END, f"無音{number_str}：{pause[1]:.2f}秒\n")
        
            self.pause_index_to_line_index[original_index] = line_index 
            line_index += 1

            if sort_by_length and divider_index is not None and i == divider_index - 1:
                self.pause_text.insert(tk.END, "――――――――――――――――――――――――――――――――\n", "divider")
                self.pause_text.tag_config("divider", foreground="#FF0044")
                line_index += 1  
    
        line_count = len(self.sorted_pauses_for_display) + (1 if divider_index is not None else 0)
        self.pause_text.config(height=line_count)
    
        self.pause_text.config(state='disabled')

    def get_pauses_in_interval(self, start_sample, end_sample):
        sr = self.sr
        pauses_in_range = []
        for i in range(len(self.intervals) - 1):
            prev_end = self.intervals[i][1]
            next_start = self.intervals[i + 1][0]
            pause_start_sec = prev_end / sr
            pause_end_sec = next_start / sr
            if prev_end >= start_sample and next_start <= end_sample:
                pauses_in_range.append((pause_start_sec, pause_end_sec))
        return pauses_in_range
    
    def on_graph_click(self, event):
        if event.inaxes is None:
            return
    
        clicked_time = event.xdata
        for i, (span, pause_index, start, end) in enumerate(self.span_artists):
            if start <= clicked_time <= end:
                for s, _, _, _ in self.span_artists:
                    s.set_alpha(0.3)
    
                span.set_alpha(0.7)
                self.selected_span_index = pause_index

                self.highlight_pause_list(pause_index)
    
                self.canvas.draw()
                break
    
    def highlight_pause_list(self, pause_index):
        self.pause_text.config(state='normal')
        self.pause_text.tag_remove('selected', '1.0', tk.END)
    
        if pause_index in self.pause_index_to_line_index:
            line_num = self.pause_index_to_line_index[pause_index]
            start = f"{line_num + 1}.0"
            end = f"{line_num + 1}.end"
            self.pause_text.tag_add('selected', start, end)
            self.pause_text.tag_config('selected', background='#8BF6EC')
    
        self.pause_text.config(state='disabled')
    
  
if __name__ == "__main__":
    root = tk.Tk()
    app = PitchViewer(root)
    root.deiconify() 
    root.mainloop()
    plt.close('all')



#ソースコードはプログラムの透明性確保のため公開しています。
#ソースコードを改変して使用した場合、その動作保証はできかねます。
#また、無断での再配布・商用利用は禁止します。

#もし何かお困りのことがありましたら、開発者（kitaohiori@gmail.com）にご連絡ください。