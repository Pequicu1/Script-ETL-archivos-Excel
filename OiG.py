import os
import subprocess
import pyautogui
import pygetwindow
import time


def ObrirGuardar(file, path):
    # Abrir los archivos de Excel uno por uno
    print(file)
    file_path = os.path.join(path, file)

    print(file_path)
    subprocess.Popen([file_path], shell=True)

    time.sleep(1.5)

    pyautogui.hotkey("ctrl", "g")
    time.sleep(1)
    pyautogui.hotkey("alt", "f4")
    time.sleep(1)
