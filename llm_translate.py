import os
import time
import traceback
import threading
import keyboard
import openai
import pystray
import win32clipboard
import win32api
from PIL import Image, ImageDraw
import tkinter as tk
import tkinter.font as tkFont
import queue
import ollama
import win32gui
import ctypes

# API プロバイダーの切り替え（"openai" "deepseek" "ollama"）
API_PROVIDER = os.getenv("LLM_TRANSLATE_API_PROVIDER", "openai")
print(f"API_PROVIDER: {API_PROVIDER}")

if API_PROVIDER == "openai":
    # OpenAI APIキー（環境変数などから取得）
    api_key = os.getenv("LLM_TRANSLATE_OPENAI_API_KEY", None)
    base_url = None
    model = "gpt-3.5-turbo"
elif API_PROVIDER == "deepseek":
    # Deep Seek APIキー（環境変数などから取得）
    api_key = os.getenv("LLM_TRANSLATE_DEEPSEEK_API_KEY", None)
    base_url = "https://api.deepseek.com"
    model = "deepseek-chat"
elif API_PROVIDER == "ollama":
    # Ollama APIキー（環境変数などから取得）
    model = os.getenv("LLM_TRANSLATE_OLLAMA_MODEL", "deepseek-r1")

if API_PROVIDER in ["openai", "deepseek"]:
    client = openai.OpenAI(api_key=api_key, base_url=base_url)
elif API_PROVIDER == "ollama":
    client = ollama


# 利用可能な翻訳先言語（ISO 639-1コード）と表示名
TARGET_LANGUAGES = ["ja", "en", "zh"]
LANGUAGE_NAMES = {
    "ja": "日本語",
    "en": "英語",
    "zh": "中国語"
}
current_target_index = 0  # 初期は "ja"
src_lang = ""

# Tkinterスレッドとの通信用キュー
# "create"メッセージはポップアップ作成、"update"メッセージは翻訳結果更新用
translation_queue = queue.Queue()

# キーボードインプット
# 定数の定義
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
VK_MENU = 0x12  # Alt キー

# KEYBDINPUT 構造体の定義
class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_ushort),
        ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))
    ]

# INPUT 構造体の定義
class INPUT(ctypes.Structure):
    _fields_ = [
        ("type", ctypes.c_ulong),
        ("ki", KEYBDINPUT)
    ]

########################################################################
# 翻訳用関数（API プロバイダーによって分岐）
########################################################################
def chat_with_openai(model, system_role, prompt):
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_role},
            {"role": "user", "content": prompt},
        ],
        temperature=0,
    )
    return response.choices[0].message.content.strip()

def chat_with_ollama(model, system_role, prompt):
    response = ollama.chat(model=model, messages=[
        {"role": "system", "content": system_role},
        {"role": "user", "content": prompt},
    ])
    res = response['message']['content']
    if "</think>" in res:
        return res.split("</think>")[1].strip()
    else:
        return res

def detect_language(text: str) -> str:
    """OpenAI を用いてテキストの言語(2文字コード)を検出する"""
    try:
        system_role = "You are a language detection tool."
        prompt = (
            "You are a language detection tool. "
            "Output only the two-letter ISO language code (e.g. 'en', 'ja', 'zh'), "
            "no extra text:\n\n"
            f"Text: {text}"
        )
        if API_PROVIDER in ["openai", "deepseek"]:
            detected = chat_with_openai(model, system_role, prompt).lower()
        elif API_PROVIDER == "ollama":
            detected = chat_with_ollama(model, system_role, prompt).lower()
        if len(detected) > 2:
            detected = detected[:2]
        return detected
    except Exception:
        return "en"  # エラー時は英語

def translate_text(text: str, src_lang: str, tgt_lang: str) -> str:
    """OpenAI を用いて、src_lang から tgt_lang へテキストを翻訳する"""
    system_role = "You are a helpful translation assistant."
    prompt = (
        f"Translate the following text from {src_lang} to {tgt_lang}. "
        "Output only the translated text:\n\n"
        f"{text}"
    )
    if API_PROVIDER in ["openai", "deepseek"]:
        translated = chat_with_openai(model, system_role, prompt)
    elif API_PROVIDER == "ollama":
        translated = chat_with_ollama(model, system_role, prompt)
    return translated

########################################################################
# クリップボード操作
########################################################################
def get_clipboard_text() -> str:
    """クリップボードからテキストを取得する"""
    text = ""
    try:
        win32clipboard.OpenClipboard()
        text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
    except Exception:
        pass
    return text

def set_clipboard_text(text: str):
    """クリップボードにテキストをセットする"""
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
    except Exception:
        pass

########################################################################
# 翻訳処理とポップアップ表示の起動
########################################################################
last_copy_time = 0
double_press_interval = 0.5  # 秒

def bring_window_to_front(window):
    window.lift()
    window.focus_force()

def on_hotkey_press(_e=None):
    """
    Ctrl+C が押されたときのコールバック  
    連打されたら、クリップボードの内容を取得し、既存のポップアップがあれば前面に持ってくる。  
    その後、別スレッドで翻訳処理を行い、結果をポップアップ更新用メッセージとして送る。
    """
    global last_copy_time, current_target_index, current_popup
    now = time.time()
    if (now - last_copy_time) < double_press_interval:
        original_text = get_clipboard_text().strip()
        if not original_text:
            last_copy_time = now
            return
        # もし既にポップアップが存在していれば、その入力エリアを更新して前面に持ってくる
        if current_popup is not None:
            current_popup.text_original.config(state="normal")
            current_popup.text_original.delete("1.0", "end")
            current_popup.text_original.insert("1.0", original_text)
            current_popup.text_original.config(state="normal")
            # 翻訳後テキストをLoadingにする
            current_popup.text_translated.config(state="normal")
            current_popup.text_translated.delete("1.0", "end")
            current_popup.text_translated.insert("1.0", "Loading...")
            current_popup.text_translated.config(state="disabled")
            bring_window_to_front(current_popup.popup)
        else:
            translation_queue.put({"type": "create", "original_text": original_text})

        # 翻訳処理を別スレッドで実行
        threading.Thread(target=perform_translation, args=(original_text,), daemon=True).start()
    last_copy_time = now

def perform_translation(original_text):
    """
    言語検出と翻訳を実行し、結果を Tkinter 側へ更新メッセージとして送る  
    エラー発生時は、エラー内容とスタックトレースを送信する。
    """
    global current_target_index, src_lang
    try:
        detected_src_lang = detect_language(original_text)
        tgt_lang = TARGET_LANGUAGES[current_target_index]
        # 入力テキストの言語が対象言語と同じなら、最初は対象言語を変更、以降は翻訳ペアを入れ替える
        if detected_src_lang.startswith(tgt_lang):
            if src_lang == "":
                current_target_index = (current_target_index + 1) % len(TARGET_LANGUAGES)
                tgt_lang = TARGET_LANGUAGES[current_target_index]
            else:
                tgt_lang = src_lang
        src_lang = detected_src_lang
        translated_text = translate_text(original_text, src_lang, tgt_lang)
        set_clipboard_text(translated_text)
        translation_queue.put({
            "type": "update",
            "src_lang": src_lang,
            "tgt_lang": tgt_lang,
            "translated_text": translated_text,
            "error": False
        })
    except Exception as e:
        error_message = str(e) + "\n" + traceback.format_exc()
        translation_queue.put({
            "type": "update",
            "src_lang": "",
            "tgt_lang": "",
            "translated_text": error_message,
            "error": True
        })

########################################################################
# TranslationPopup クラス (Tkinter のポップアップ)
########################################################################
current_popup = None  # 現在表示中のポップアップ

class TranslationPopup:
    def __init__(self, parent, original_text):
        self.parent = parent
        self.popup = tk.Toplevel(parent)
        self.popup.title("Translation")
        # 生成時に一度前面に持ってくるが、常に最前面にはしない
        self.popup.lift()
        self.popup.resizable(True, True)
        self.popup.protocol("WM_DELETE_WINDOW", self.close)
        self.popup.geometry("600x400")
        self.popup.bind("<Configure>", self.on_configure)

        self.label_font = tkFont.Font(family="Arial", size=12, weight="bold")
        self.text_font = tkFont.Font(family="Arial", size=12)

        self.main_frame = tk.Frame(self.popup, bg="white", bd=2, relief="solid")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.popup.grid_rowconfigure(0, weight=1)
        self.popup.grid_columnconfigure(0, weight=1)

        self.left_frame = tk.Frame(self.main_frame, bg="white")
        self.right_frame = tk.Frame(self.main_frame, bg="white")
        self.left_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.right_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)

        # --- 左側：翻訳元エリア（入力可能） ---
        self.label_original_title = tk.Label(self.left_frame, text="Original (??)", bg="white", font=self.label_font)
        self.label_original_title.grid(row=0, column=0, sticky="w")
        self.left_text_frame = tk.Frame(self.left_frame, bg="white")
        self.left_text_frame.grid(row=1, column=0, sticky="nsew")
        self.left_frame.grid_rowconfigure(1, weight=1)
        self.left_frame.grid_columnconfigure(0, weight=1)
        self.left_text_frame.grid_rowconfigure(0, weight=1)
        self.left_text_frame.grid_columnconfigure(0, weight=1)
        self.left_text_frame.grid_columnconfigure(1, weight=0)
        self.text_original = tk.Text(self.left_text_frame, wrap="word", font=self.text_font, bg="lightgrey", width=28)
        self.text_original.grid(row=0, column=0, sticky="nsew")
        self.orig_scrollbar = tk.Scrollbar(self.left_text_frame, orient="vertical", command=self.text_original.yview)
        self.orig_scrollbar.grid(row=0, column=1, sticky="ns")
        self.text_original.config(yscrollcommand=self.orig_scrollbar.set)
        self.text_original.insert("1.0", original_text)
        self.text_original.bind("<Shift-Return>", self.on_shift_enter)
        self.btn_copy_original = tk.Button(self.left_frame, text="📋", command=lambda: set_clipboard_text(self.text_original.get("1.0", "end").strip()))
        self.btn_copy_original.grid(row=2, column=0, sticky="w", pady=(5,0))
        self.btn_translate = tk.Button(self.left_frame, text="翻訳する", command=self.start_translation_from_input)
        self.btn_translate.grid(row=3, column=0, sticky="w", pady=(5,0))

        # --- 右側：翻訳結果エリア ---
        self.label_translated_title = tk.Label(self.right_frame, text="Translation (??)", bg="white", font=self.label_font)
        self.label_translated_title.grid(row=0, column=0, sticky="w")
        self.right_text_frame = tk.Frame(self.right_frame, bg="white")
        self.right_text_frame.grid(row=1, column=0, sticky="nsew")
        self.right_frame.grid_rowconfigure(1, weight=1)
        self.right_frame.grid_columnconfigure(0, weight=1)
        self.right_text_frame.grid_rowconfigure(0, weight=1)
        self.right_text_frame.grid_columnconfigure(0, weight=1)
        self.right_text_frame.grid_columnconfigure(1, weight=0)
        self.text_translated = tk.Text(self.right_text_frame, wrap="word", font=self.text_font, bg="lightgrey", width=28)
        self.text_translated.grid(row=0, column=0, sticky="nsew")
        self.trans_scrollbar = tk.Scrollbar(self.right_text_frame, orient="vertical", command=self.text_translated.yview)
        self.trans_scrollbar.grid(row=0, column=1, sticky="ns")
        self.text_translated.config(yscrollcommand=self.trans_scrollbar.set)
        self.text_translated.insert("1.0", "Loading...")
        self.text_translated.config(state="disabled")
        self.btn_copy_translated = tk.Button(self.right_frame, text="📋", command=lambda: set_clipboard_text(self.get_translated_text()))
        self.btn_copy_translated.grid(row=2, column=0, sticky="w", pady=(5,0))

    def on_configure(self, event):
        base_width = 600
        base_height = 400
        current_width = self.popup.winfo_width()
        current_height = self.popup.winfo_height()
        scale = min(current_width / base_width, current_height / base_height)
        new_size = max(12, int(12 * scale))
        self.label_font.config(size=new_size)
        self.text_font.config(size=new_size)
        self.text_original.config(font=self.text_font)
        self.text_translated.config(font=self.text_font)
        self.label_original_title.config(font=self.label_font)
        self.label_translated_title.config(font=self.label_font)

    def update(self, src_lang, tgt_lang, translated_text, error=False):
        if error:
            self.text_translated.config(fg="red")
            self.label_translated_title.config(text="Error")
        else:
            self.text_translated.config(fg="black")
            original_lang_name = LANGUAGE_NAMES.get(src_lang, src_lang.upper())
            translation_lang_name = LANGUAGE_NAMES.get(tgt_lang, tgt_lang.upper())
            self.label_original_title.config(text=f"Original ({original_lang_name})")
            self.label_translated_title.config(text=f"Translation ({translation_lang_name})")
        self.text_translated.config(state="normal")
        self.text_translated.delete("1.0", "end")
        self.text_translated.insert("1.0", translated_text)
        self.text_translated.config(state="disabled")
        # ウィンドウを前面に持ってくる
        try:
            bring_window_to_front(self.popup)
        except Exception:
            self.popup.lift()
            self.popup.focus_force()
        self.popup.attributes("-topmost", False)

    def on_shift_enter(self, event):
        self.start_translation_from_input()
        return "break"

    def start_translation_from_input(self):
        original_text = self.text_original.get("1.0", "end").strip()
        if original_text:
            self.text_translated.config(state="normal")
            self.text_translated.delete("1.0", "end")
            self.text_translated.insert("1.0", "Loading...")
            self.text_translated.config(state="disabled")
            threading.Thread(target=perform_translation, args=(original_text,), daemon=True).start()

    def get_translated_text(self):
        return self.text_translated.get("1.0", "end").strip()

    def close(self):
        self.popup.destroy()
        global current_popup
        current_popup = None

def bring_window_to_front(window):
    hwnd = window.winfo_id()
    win32api.keybd_event(0x12, 0, 0, 0)  # ALT キー押下
    try:
        # win32gui.ShowWindow(hwnd, 9)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.1)
    except Exception:
        window.lift()
        window.focus_force()
    win32api.keybd_event(0x12, 0, 2, 0)  # ALT キー解放

    


########################################################################
# Tkinterメインループ
########################################################################

def tkinter_mainloop():
    global current_popup, root
    root = tk.Tk()
    root.withdraw()

    def poll_queue():
        global current_popup
        try:
            while True:
                message = translation_queue.get_nowait()
                if isinstance(message, dict):
                    if message.get("type") == "create":
                        if current_popup is not None:
                            # 更新：既存のポップアップがあれば翻訳元テキストを更新する
                            current_popup.text_original.config(state="normal")
                            current_popup.text_original.delete("1.0", "end")
                            current_popup.text_original.insert("1.0", message["original_text"])
                            current_popup.text_original.config(state="normal")
                        else:
                            current_popup = TranslationPopup(root, message["original_text"])

                    elif message.get("type") == "update":
                        if current_popup is not None:
                            current_popup.update(
                                message.get("src_lang", ""),
                                message.get("tgt_lang", ""),
                                message.get("translated_text", ""),
                                message.get("error", False)
                            )
                elif isinstance(message, str) and message.startswith("[Language Info]"):
                    create_info_popup(root, message.replace("[Language Info] ", ""), duration=1500)
        except queue.Empty:
            pass
        root.after(100, poll_queue)

    poll_queue()
    root.mainloop()

def create_info_popup(parent, message, duration=1500):
    popup = tk.Toplevel(parent)
    popup.overrideredirect(True)
    width = 300
    height = 50
    mouse_x, mouse_y = win32api.GetCursorPos()
    x = mouse_x - width // 2
    y = mouse_y - height - 5
    popup.geometry(f"{width}x{height}+{x}+{y}")
    frame = tk.Frame(popup, bg="lightyellow", bd=2, relief="solid")
    frame.pack(fill="both", expand=True)
    label = tk.Label(frame, text=message, bg="lightyellow", wraplength=width-20)
    label.pack(padx=10, pady=10)
    popup.after(duration, popup.destroy)

########################################################################
# pystrayアイコン設定
########################################################################

def tray_icon_setup():
    icon_size = (64, 64)
    image = Image.new("RGB", icon_size, (0, 128, 255))
    draw = ImageDraw.Draw(image)
    draw.text((10, 20), "Tr", fill=(255, 255, 255))

    icon = pystray.Icon("OpenAI Translator", image, "OpenAI Translator")

    def on_quit(icon, _item):
        icon.stop()

    def switch_lang(icon, _item):
        global current_target_index
        current_target_index = (current_target_index + 1) % len(TARGET_LANGUAGES)
        new_lang = TARGET_LANGUAGES[current_target_index]
        icon.title = f"Target: {LANGUAGE_NAMES.get(new_lang, new_lang)}"
        translation_queue.put(f"[Language Info] Translation target switched to: {LANGUAGE_NAMES.get(new_lang, new_lang)}")

    icon.menu = pystray.Menu(
        pystray.MenuItem("Switch Language (ja/en/zh)", switch_lang),
        pystray.MenuItem("Quit", on_quit)
    )
    return icon

def run_tray_icon():
    icon = tray_icon_setup()
    icon.run()

########################################################################
# メイン処理
########################################################################

def main():
    threading.Thread(target=tkinter_mainloop, daemon=True).start()
    keyboard.add_hotkey("ctrl+c", on_hotkey_press)
    run_tray_icon()

if __name__ == "__main__":
    main()
