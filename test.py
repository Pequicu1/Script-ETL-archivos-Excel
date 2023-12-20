import os
import subprocess
import pyautogui
import pygetwindow as gw
import time


# wb = openpyxl.load_workbook(
#     "Mila\\Colombia\\Almacenes Exito SA.xlsm", data_only=True, keep_vba=True)

# ws = wb["Valoración por Flujos de Caja"]

# print(ws['C35'].value)

# Get a list of all open window titles
window_titles = gw.getAllTitles()
excel_wind = gw.getWindowsWithTitle('Acerias Paz del Rio SA')[0]
# Print the titles of all open windows
excel_wind.close()
