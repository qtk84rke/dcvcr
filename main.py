import os
import sys
import time
import hashlib
import threading
import subprocess

from win32con import *
from win32gui import *
from win32process import *

WAIT_SEC = 0.5
WINDOW_NAME = 'VOICEROID2'

def talk(text: str):
    output_dir = './output/'
    try:
        os.mkdir(output_dir)
    except:
        pass

    output_file = output_dir + hashlib.md5(text.encode('utf-8')).hexdigest() + '.mp3'
    if os.path.exists(output_file):
        return output_file

    temp_file = 'temp.wav'
    while True:
        if os.path.exists(output_file):
            time.sleep(WAIT_SEC)
        else:
            break

    while True:
        window = FindWindow(None, WINDOW_NAME) or FindWindow(None, WINDOW_NAME + '*')
        if window == 0:
            subprocess.Popen([r'E:\applications\AHS\VOICEROID2\VoiceroidEditor.exe'])
            time.sleep(3 * WAIT_SEC)
        else:
            break

    while True:
        error_dialog = FindWindow(None, 'エラー') or FindWindow(None, '注意') or FindWindow(None, '音声ファイルの保存')
        if error_dialog:
            SendMessage(error_dialog, WM_CLOSE, 0, 0)
            time.sleep(WAIT_SEC)
        else:
            break

    SetWindowPos(window, HWND_TOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW | SWP_NOMOVE | SWP_NOSIZE)

    def enum_dialog_callback(hwnd, param):
        class_name = GetClassName(hwnd)
        win_text = GetWindowText(hwnd)

        if class_name.count('Edit'):
            SendMessage(hwnd, WM_SETTEXT, 0, temp_file)

        if win_text.count('保存'):
            SendMessage(hwnd, WM_LBUTTONDOWN, MK_LBUTTON, 0)
            SendMessage(hwnd, WM_LBUTTONUP, 0, 0)

    def save():
        time.sleep(WAIT_SEC)

        dialog = FindWindow(None, '音声ファイルの保存')
        if dialog:
            EnumChildWindows(dialog, enum_dialog_callback, None)
            return

        save()

    def enum_callback(hwnd, param):
        class_name = GetClassName(hwnd)
        win_text = GetWindowText(hwnd)

        if class_name.count('RichEdit20W'):
            SendMessage(hwnd, WM_SETTEXT, 0, text)

        if win_text.count('音声保存'):
            ShowWindow(window, SW_SHOWNORMAL)

            threading.Thread(target=save).start()

            SendMessage(hwnd, WM_LBUTTONDOWN, MK_LBUTTON, 0)
            SendMessage(hwnd, WM_LBUTTONUP, 0, 0)

    EnumChildWindows(window, enum_callback, None)

    while True:
        if FindWindow(None, '音声保存'):
            time.sleep(WAIT_SEC)
        else:
            break

    #subprocess.run(['ffmpeg', '-i', temp_file, '-acodec', 'libmp3lame', '-ab', '128k', '-ac', '2', '-ar', '44100', output_file])

    try:
        os.remove(temp_file)
        os.remove(temp_file.replace('wav', 'txt'))
    except:
        pass

    return output_file

if __name__ == '__main__':
    if len(sys.argv) > 1:
        print(talk(sys.argv[1]))
    else:
        print(talk('パイソンからの読み上げテストです'))
