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

if __name__ == '__main__':
    talk('テストです')
