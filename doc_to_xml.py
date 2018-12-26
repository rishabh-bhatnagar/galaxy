from pyautogui import *
from os import listdir
from time import sleep as wait

n_files = 0
for i in listdir('C:\\Users\\rishabh\Desktop\\rpa projects\\rpae_project\\all_opf\\delete_docs'):
    n_files += 1

ctrl = 'ctrl'
alt = 'alt'
enter = 'enter'
tab = 'tab'
hotkey(alt, tab)
for i in range(n_files):
    press('down')
    press(enter)
    wait(2)
    hotkey(alt, 'f')
    wait(0.3)
    hotkey(alt, 'a')
    wait(0.3)
    hotkey(alt, '2')
    wait(0.3)
    typewrite(str(i))
    press(tab)
    press('down')
    press('up', interval=1)
    for j in range(14):
        press('down')
    press(enter)
    wait(0.3)
    press(enter)
    wait(0.7)
    hotkey(alt,'f4')
    wait(0.6)


