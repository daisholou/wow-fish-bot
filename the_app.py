# -*- coding: utf-8 -*-

import pyautogui
# import win32api
import win32gui
# import win32ui
import win32con
from threading import Thread
import os
import numpy as np
import tkinter as tk
# import tkinter.font as tkFont
import tkinter.messagebox as tkmsgbox
from tkinter import scrolledtext
# import hashlib
import time
import cv2
import sounddevice as sd
import pyaudio
import audioop
import math
from collections import deque
import wave
from skimage.metrics import structural_similarity
import imutils


class APP:

    def __init__(self, init_window):
        self.name = None
        self.log_data_Text = None
        self.log_label = None
        self.online_wnd_list = None
        self.online_label = None
        self.show_button = None
        self.refresh_button = None
        self.down_button = None
        self.up_button = None
        self.free_wnd_list = None
        self.free_label = None
        self.log_line_num = 0
        self.init_window = init_window
        self.free_hwnd = []
        self.free_name = []
        self.online_hwnd = []
        self.online_name = []
        self.wows = {}
        self.disk_lock = True

    # 设置窗口
    def set_init_window(self):
        self.init_window.title("WoW Fishing V1.00")  # 窗口名
        width = 600

        height = 600
        screenwidth = self.init_window.winfo_screenwidth()
        screenheight = self.init_window.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.init_window.geometry(alignstr)
        self.init_window.resizable(width=False, height=False)

        self.free_label = tk.Label(self.init_window, text="自由窗口")
        self.free_label.place(x=20, y=10, width=70, height=20)

        self.free_wnd_list = tk.Listbox(self.init_window, selectmode=tk.MULTIPLE)
        self.free_wnd_list.place(x=20, y=40, width=560, height=60)
        self.refresh()

        self.up_button = tk.Button(self.init_window, text="移除", bg="lightblue", width=10, command=self.remove)
        self.up_button.place(x=400, y=120, width=70, height=25)

        self.down_button = tk.Button(self.init_window, text="加入", bg="lightblue", width=10, command=self.join)
        self.down_button.place(x=300, y=120, width=70, height=25)

        self.refresh_button = tk.Button(self.init_window, text="刷新", bg="lightblue", width=10,
                                        command=self.refresh)
        self.refresh_button.place(x=200, y=120, width=70, height=25)

        self.show_button = tk.Button(self.init_window, text="显示窗口", bg="lightblue", width=10,
                                     command=self.show_window)
        self.show_button.place(x=100, y=120, width=70, height=25)

        self.online_label = tk.Label(self.init_window, text="托管窗口")
        self.online_label.place(x=20, y=160, width=70, height=20)

        self.online_wnd_list = tk.Listbox(self.init_window, selectmode=tk.MULTIPLE)
        self.online_wnd_list.place(x=20, y=190, width=560, height=60)

        # 日志框
        self.log_label = tk.Label(self.init_window, text="日志")
        self.log_label.place(x=20, y=270, width=70, height=20)

        self.log_data_Text = tk.scrolledtext.ScrolledText(self.init_window)
        self.log_data_Text.place(x=20, y=300, width=560, height=240)

    # 删除
    def remove(self):
        wnd = self.online_wnd_list.curselection()
        if wnd:
            for e in wnd:
                hwnd = self.online_hwnd[e]
                print(hwnd)
                self.free_hwnd.append(hwnd)
                self.free_name.append(self.wows[hwnd].name)
                self.free_wnd_list.insert('end', self.free_name[-1])
                self.wows[hwnd].working = False
                self.write_log_to_text(self.free_name[-1] + ' stoped')

            for e in reversed(wnd):
                self.online_hwnd.__delitem__(e)
                self.online_name.__delitem__(e)
                self.online_wnd_list.delete(e)

            print(self.free_hwnd)
            print(self.online_hwnd)

    # 添加
    def join(self):
        wnd = self.free_wnd_list.curselection()
        if wnd:
            for e in wnd:
                hwnd = self.free_hwnd[e]
                print(hwnd)
                self.online_hwnd.append(hwnd)
                self.online_name.append(self.wows[hwnd].name)
                self.online_wnd_list.insert('end', self.online_name[-1])
                self.write_log_to_text(self.online_name[-1] + ' start auto fishing')
                self.wows[hwnd].working = True
                self.wows[hwnd].run()
                # self.get_screen(hwnd)

            for e in reversed(wnd):
                self.free_hwnd.__delitem__(e)
                self.free_name.__delitem__(e)
                self.free_wnd_list.delete(e)

            print(self.free_hwnd)
            print(self.online_hwnd)

    # 刷新数据
    def refresh(self):
        wnd = self.find_window(self.name)

        if len(self.wows) > 0:
            for i in self.wows.keys():
                if i not in wnd:
                    self.wows[i].__del__()
                    self.wows.__delitem__(i)

        if len(self.free_hwnd) > 0:
            for i in range(len(self.free_hwnd) - 1, -1, -1):
                if self.free_hwnd[i] not in wnd:
                    self.free_hwnd.__delitem__(i)
                    self.free_name.__delitem__(i)
                    self.free_wnd_list.delete(i)
        if len(self.online_hwnd) > 0:
            for i in range(len(self.online_hwnd) - 1, -1, -1):
                if self.online_hwnd[i] not in wnd:
                    self.online_hwnd.__delitem__(i)
                    self.online_name.__delitem__(i)
                    self.online_wnd_list.delete(i)

        for hwnd in wnd:
            if hwnd not in self.wows.keys():
                self.wows[hwnd] = WOW(hwnd, self)
                print(self.wows)
            if hwnd not in self.free_hwnd and hwnd not in self.online_hwnd:
                self.free_hwnd.append(hwnd)
                self.free_name.append(self.wows[hwnd].name)
                self.free_wnd_list.insert('end', self.free_name[-1])

        print(self.free_hwnd)
        print(self.online_hwnd)

    # 获取所有窗口
    def find_window(self, targettitle=None):
        fwnd = []
        hWndList = []
        win32gui.EnumWindows(lambda hWnd, param: param.append(hWnd), hWndList)
        # print(hWndList)
        if targettitle is None:
            for hwnd in hWndList:
                title = win32gui.GetWindowText(hwnd)
                if len(title) > 0:
                    fwnd.append(hwnd)
            # print(fwnd)
            return fwnd[0:15]
        else:
            for hwnd in hWndList:
                title = win32gui.GetWindowText(hwnd)
                if title.find(targettitle) >= 0:
                    fwnd.append(hwnd)
            # print(fwnd)
            return fwnd

    def get_window_name(self, hwnd):
        return win32gui.GetWindowText(hwnd)

    def show_window(self):
        wnd = self.free_wnd_list.curselection()
        if wnd:
            for e in wnd:
                hwnd = self.free_hwnd[e]
                print(hwnd)
                win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)
                time.sleep(0.5)

    def get_window_status(self, hwnd):
        status = '未知'
        if self.find_specify_picture(os.getcwd() + "\\wow_login_in.png"):
            status = '人物界面'
        if self.find_specify_picture(os.getcwd() + "\\wow_disconnect.png") or self.find_specify_picture(
                os.getcwd() + "\\wow_disconnect_too.png"):
            status = '断开连接'
        if self.find_specify_picture(os.getcwd() + "\\wow_reconnect.png"):
            status = '重新连接'

    # 获取当前时间
    def get_current_time(self):
        current_time = time.strftime('%m-%d %H:%M:%S', time.localtime(time.time()))
        return current_time

    # 日志动态打印
    def write_log_to_text(self, logmsg):

        current_time = self.get_current_time()
        log_msg_in = str(current_time) + " " + str(logmsg) + "\n"  # 换行
        self.log_data_Text.insert(tk.END, log_msg_in)
        self.log_line_num += 1

    def on_closing(self):
        if tkmsgbox.askokcancel(u"退出", u"确定退出程序吗？"):
            self.init_window.destroy()


class WOW:

    def __init__(self, hwnd, app):
        super(WOW, self).__init__()
        self.left = 0
        self.top = 0
        self.right = 0
        self.bot = 0
        self.width = 0
        self.height = 0
        self.hwnd = hwnd
        self.name = str(hwnd) + ' ' + win32gui.GetWindowText(hwnd)
        self.working = False
        self.need_bait = False
        self.threshold = 0.5
        self.gap = 0.5  # 每1秒检测一次

        win32gui.SetForegroundWindow(self.hwnd)
        win32gui.ShowWindow(self.hwnd, win32con.SW_SHOW)
        self.left, self.top, self.right, self.bot = win32gui.GetWindowRect(self.hwnd)
        self.width = self.right - self.left
        self.height = self.bot - self.top

    def run(self):

        t = Thread(target=self.fishing)
        t.start()

    def fishing(self):
        bait_time = 0
        fish_time = 0
        last_x = 0
        last_y = 0
        is_bait = False
        is_fishing = False
        app.write_log_to_text(self.name + ' start fishing')
        while self.working:

            if not is_fishing:
                find_x = 0
                find_y = 0
                while find_x + find_y == 0:
                    pyautogui.press('1')  # 下竿
                    app.write_log_to_text(self.name + ' Lower fishing rod')
                    fish_time = time.time()
                    time.sleep(2)
                    find_x, find_y = self.get_float()
                is_fishing = True
            time.sleep(self.gap)
            if time.time() - fish_time < 20 and is_fishing:
                x, y = self.get_float()
                if x + y > 0:
                    if abs(find_x - x) < 2 and y - find_y > 2:
                        pyautogui.moveTo(find_x + self.left, find_y + self.top, 1)  # 收杆
                        # pyautogui.keyDown("shiftleft")
                        pyautogui.click(button='right')
                        # pyautogui.keyUp("shiftleft")
                        app.write_log_to_text(self.name + ' catch U!!')
                        print("fishing time :", time.time() - fish_time)
                        is_fishing = False
                        time.sleep(3)
                    last_x = x
                    last_y = y
                else:
                    if last_x + last_y > 0:
                        pyautogui.moveTo(find_x + self.left, find_y + self.top, 1)  # 收杆
                        # pyautogui.keyDown("shiftleft")
                        pyautogui.click(button='right')
                        # pyautogui.keyUp("shiftle
                        # ft")
                        app.write_log_to_text(self.name + ' catch U!!')
                        print("fishing time :", time.time() - fish_time)
                        time.sleep(3)
                        is_fishing = False
            else:
                app.write_log_to_text(self.name + ' nothing happens')
                print("fishing time :", time.time() - fish_time)
                is_fishing = False
            if time.time() - bait_time > 600:
                is_bait = False
            if self.need_bait and not is_bait and not is_fishing:
                pyautogui.press('2')  # 上饵
                bait_time = time.time()
                app.write_log_to_text(self.name + ' add bait')
                is_bait = True

    def find_specify_picture(self, targetpicture):
        character = pyautogui.locateOnScreen(targetpicture, confidence=0.5)  # grayscale=True)
        if (character is not None):
            imcenter = pyautogui.center(character)
            print(imcenter)
            return imcenter
        else:
            print("none posion find")
            return None

    def find_float(self, img1, img2):

        float_x, float_y, float_width, float_height = 0

        grayA = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
        grayB = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
        (score, diff) = structural_similarity(grayA, grayB, full=True)
        diff = (diff * 255).astype("uint8")
        print("SSIM:{}".format(score))

        # 找到不同点的轮廓以致于我们可以在被标识为“不同”的区域周围放置矩形
        thresh = cv2.threshold(diff, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
        cnts = cv2.findContours(thresh.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        cnts = cnts[0] if imutils.is_cv2() else cnts[1]

        # 找到一系列区域，在区域周围放置矩形
        for c in cnts:
            (x, y, w, h) = cv2.boundingRect(c)
            cv2.rectangle(img1, (x, y), (x + w, y + h), (0, 255, 0), 2)
            cv2.rectangle(img2, (x, y), (x + w, y + h), (0, 255, 0), 2)
            if 3 < w < 10 and 3 < h < 10:
                return x, y, w, h
        # 用cv2.imshow 展现最终对比之后的图片， cv2.imwrite 保存最终的结果图片
        cv2.imshow("Modified", img2)
        cv2.imwrite(r"result.bmp", img2)
        cv2.waitKey(0)

        return float_x, float_y, float_width, float_height

    def get_float(self):
        # 加载原始的rgb图像
        img_rgb = self.get_screen()
        # 创建一个原始图像的灰度版本，所有操作在灰度版本中处理，然后在RGB图像中使用相同坐标还原
        img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
        # self.show_img(img_gray)
        # 加载将要搜索的图像模板
        template = cv2.imread('n.png', cv2.IMREAD_GRAYSCALE)
        # template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
        # height, width = template.shape[:2]
        # size = (int(width * 0.5), int(height * 0.5))
        # template = cv2.resize(template, size, interpolation=cv2.INTER_AREA)
        # self.show_img(template)
        # 记录图像模板的尺寸
        w, h = template.shape[::-1]

        res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF_NORMED)
        # res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF)
        #
        # 'cv2.TM_CCOEFF', 'cv2.TM_CCOEFF_NORMED', 'cv2.TM_CCORR',
        # 'cv2.TM_CCORR_NORMED', 'cv2.TM_SQDIFF', 'cv2.TM_SQDIFF_NORMED'
        # cv2.TM_SQDIFF, cv2.TM_SQDIFF_NORMED 是最小值
        # self.show_img(res)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        print(min_val, max_val, min_loc, max_loc)
        if max_val > self.threshold:
            print('找到的坐标', max_loc)
            coordinate = (self.width / 3 + max_loc[0] + w / 2, self.height / 2 + max_loc[1] + h / 2)  # 中心位置
            bottom_right = (max_loc[0] + w, max_loc[1] + h)  # 右下角的位置
            cv2.rectangle(img_rgb, max_loc, bottom_right, (255, 255, 0), 3)
            cv2.imwrite("img.jpg", img_rgb, [int(cv2.IMWRITE_JPEG_QUALITY), 100])
        else:
            coordinate = (0, 0)
        return coordinate

    def get_screen(self, x, y, width, height):
        if x and y and width and height:
            img = pyautogui.screenshot(region=[x, y, width, height])
        else:
            img = pyautogui.screenshot(
                region=[self.left + self.width / 3, self.top + self.height / 2, self.width / 3, self.height / 3])
        img = cv2.cvtColor(np.asarray(img), cv2.COLOR_RGB2BGR)
        return img

    def show_img(self, img):
        cv2.namedWindow('test')  # 命名窗口
        cv2.imshow("test", img)  # 显示
        cv2.waitKey(0)
        cv2.destroyAllWindows()
        return

    def listen(self):
        print('Well, now we are listening for loud sounds...')
        CHUNK = 1024  # CHUNKS of bytes to read each time from mic
        FORMAT = pyaudio.paInt16
        CHANNELS = 2
        RATE = 18000
        THRESHOLD = 1500  # The threshold intensity that defines silence
        # and noise signal (an int. lower than THRESHOLD is silence).
        SILENCE_LIMIT = 1  # Silence limit in seconds. The max ammount of seconds where
        # only silence is recorded. When this time passes the
        # recording finishes and the file is delivered.
        # Open stream
        p = pyaudio.PyAudio()
        device_count = p.get_device_count()
        print(f"device count: {device_count}")
        device_index = 2
        # get device info
        # for i in range(device_count):
        #     device_info = p.get_device_info_by_index(i)
        #     if device_info.get('name', '').find('立体声混音') != -1:
        #         device_index = device_info.get('index')
        #         print(device_info)
        #         print(device_index)
        stream = p.open(format=FORMAT,
                        channels=CHANNELS,
                        rate=RATE,
                        input=True,
                        input_device_index=device_index,
                        frames_per_buffer=CHUNK)
        cur_data = ''  # current chunk  of audio data
        rel = RATE / CHUNK
        slid_win = deque(maxlen=SILENCE_LIMIT * int(rel))

        success = False
        listening_start_time = time.time()
        record_buf = []
        while True:
            try:
                cur_data = stream.read(CHUNK)
                # print(cur_data)
                record_buf.append(cur_data)
                slid_win.append(math.sqrt(abs(audioop.avg(cur_data, 4))))
                print(time.time() - listening_start_time)
                # print(sum([x > THRESHOLD for x in slid_win]))
                if sum([x > THRESHOLD for x in slid_win]) > 0:
                    print('I heart something!')
                    success = True
                    break
                if time.time() - listening_start_time > 18:
                    print('I don\'t hear anything already 20 seconds!')
                    break
            except IOError:
                break

        # print "* Done recording: " + str(time.time() - start)
        wf = wave.open('01.wav', 'wb')  # 创建一个音频文件，名字为“01.wav"
        wf.setnchannels(CHANNELS)  # 设置声道数为2
        wf.setsampwidth(2)  # 设置采样深度为
        wf.setframerate(RATE)  # 设置采样率为16000
        # 将数据写入创建的音频文件
        wf.writeframes("".encode().join(record_buf))
        # 写完后将文件关闭
        wf.close()
        stream.close()
        p.terminate()
        return success


if __name__ == "__main__":
    main_window = tk.Tk()  # 实例化出一个父窗口2
    app = APP(main_window)
    app.name = '魔兽世界'
    # app.name = 'WPS Office'
    # 设置根窗口默认属性
    app.set_init_window()
    main_window.protocol("WM_DELETE_WINDOW", app.on_closing)
    main_window.mainloop()
