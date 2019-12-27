# -*- coding: utf-8 -*-
import webbrowser
import sys
import os
import struct
import time
#
import pyautogui
import numpy as np
import cv2
#
from win10toast import ToastNotifier
from PIL import ImageGrab
from win32gui import GetWindowText, GetForegroundWindow, GetWindowRect
from threading import Thread
from infi.systray import SysTrayIcon
import dlib


def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def app_pause(systray):
    global is_stop
    is_stop = False if is_stop is True else True
    # print ("Is Pause: " + str(is_stop))
    if is_stop is True:
        systray.update(
            hover_text=app + " - On Pause") 
    else:
        systray.update(
            hover_text=app)         


def app_destroy(systray):
    # print("Exit app")
    sys.exit()


def app_about(systray):
    # print("github.com/YECHEZ/wow-fish-bot")
    webbrowser.open('https://github.com/YECHEZ/wow-fish-bot')


if __name__ == "__main__":
    is_stop = True
    flag_exit = False
    lastx = 0
    lasty = 0
    is_block = False
    new_cast_time = 0
    recast_time = 40
    wait_mes = 0    
    app = "WoW Fish BOT by YECHEZ"
    link = "github.com/YECHEZ/wow-fish-bot"
    app_ico = resource_path('wow-fish-bot.ico')
    menu_options = (("Start/Stop", None, app_pause),
                    (link, None, app_about),)
    systray = SysTrayIcon(app_ico, app, 
                          menu_options, on_quit=app_destroy)
    systray.start()
    toaster = ToastNotifier()
    toaster.show_toast(app,
                       link,
                       icon_path=app_ico,
                       duration=5)
    detector = dlib.simple_object_detector("detector.svm")
    while flag_exit is False:
        if is_stop == False:
            if GetWindowText(GetForegroundWindow()) != "魔兽世界":
                if wait_mes == 5:
                    wait_mes = 0
                    toaster.show_toast(app,
                                       "Waiting for World of Warcraft"
                                       + " as active window",
                                       icon_path='wow-fish-bot.ico',
                                       duration=5)                  
                # print("Waiting for World of Warcraft as active window")
                systray.update(
                    hover_text=app
                    + " - Waiting for World of Warcraft as active window")
                wait_mes += 1
                time.sleep(2)
            else:
                systray.update(hover_text=app)
                rect = GetWindowRect(GetForegroundWindow())
                
                if not is_block:
                    lastx = 0
                    lasty = 0
                    pyautogui.press('1')
                    # print("Fish on !")
                    new_cast_time = time.time()
                    is_block = True
                    time.sleep(2)
                else:
                    fish_area = (0, rect[3] / 2, rect[2], rect[3])
    
                    img = ImageGrab.grab(fish_area)
                    img_np = np.array(img)

                    dets = detector(img_np, 1)
                    print("Number of faces detected: {}".format(len(dets)))
                    for i, d in enumerate(dets):
                        print("Detection {}: Left: {} Top: {} Right: {} Bottom: {}".format(
                            i, d.left(), d.top(), d.right(), d.bottom()))

                        b_x = int((d.left() + d.right()) / 2)
                        b_y = int((d.top() + d.bottom()) / 2)
                        pyautogui.moveTo(b_x, b_y + fish_area[1], 0.3)
                        pyautogui.keyDown('shiftleft')
                        pyautogui.mouseDown(button='right')
                        pyautogui.mouseUp(button='right')
                        pyautogui.keyUp('shiftleft')
                        # print("Catch !")
                        time.sleep(5)
                    
                    # show windows with mask
                    # cv2.imshow("fish_mask", mask)
                    # cv2.imshow("fish_frame", frame)
    
                    if time.time() - new_cast_time > recast_time:
                        # print("New cast if something wrong")
                        is_block = False               
            if cv2.waitKey(1) == 27:
                break
        else:
            # print("Pause")
            systray.update(hover_text=app + " - On Pause")   
            time.sleep(2)
