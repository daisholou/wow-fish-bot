# -*- coding: utf-8 -*-

import pyautogui
import win32api
import win32gui
# import win32ui
import win32con
from threading import Thread
import time
import os
import numpy as np
import tkinter as tk
# import tkinter.font as tkFont
import tkinter.messagebox as tkmsgbox
import hashlib
import time
import cv2
import win32ui
from PIL import ImageGrab

LOG_LINE_NUM = 0


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

        self.log_data_Text = tk.Text(self.init_window)
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
                self.write_log_to_Text(self.free_name[-1] + ' stoped')

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
                self.write_log_to_Text(self.online_name[-1] + ' start auto fishing')
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
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
    def write_log_to_Text(self, logmsg):
        global LOG_LINE_NUM
        current_time = self.get_current_time()
        logmsg_in = str(current_time) + " " + str(logmsg) + "\n"  # 换行
        if LOG_LINE_NUM <= 7:
            self.log_data_Text.insert(tk.END, logmsg_in)
            LOG_LINE_NUM = LOG_LINE_NUM + 1
        else:
            self.log_data_Text.delete(1.0, 2.0)
            self.log_data_Text.insert(tk.END, logmsg_in)

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
        self.gap = 0.5  # 每0.5秒检测一次

    def run(self):

        t = Thread(target=self.fishing)
        t.start()

    def fishing(self):
        bait_time = 0
        last_x = 0
        last_y = 0
        while self.working:
            app.write_log_to_Text(self.name + ' start fishing')
            is_bait = False
            x, y = self.get_float()
            print(x, y)
            if x + y > 0:
                if last_y != y and abs(last_x - x) < 1 and abs(last_y - y) < 3:
                    pyautogui.moveTo(x+self.left, y+self.top, 0.3)  # 收杆
                    # pyautogui.keyDown("shiftleft")
                    pyautogui.click(button='right')
                    # pyautogui.keyUp("shiftleft")
                last_x = x
                last_y = y
                time.sleep(self.gap)
            else:
                if time.localtime(time.time()) - bait_time > 600:
                    is_bait = False
                if not is_bait:
                    pyautogui.press('2')  # 上饵
                    bait_time = time.localtime(time.time())
                    app.write_log_to_Text(self.name + ' add bait')
                    is_bait = True
                pyautogui.press('1')  # 下竿
                app.write_log_to_Text(self.name + ' Lower fishing rod')

    def find_specify_picture(self, targetpicture):
        character = pyautogui.locateOnScreen(targetpicture, confidence=0.5)  # grayscale=True)
        if (character is not None):
            imcenter = pyautogui.center(character)
            print(imcenter)
            return imcenter
        else:
            print("none posion find")
            return None

    def find_float(self):

        # img = self.get_screen()
        # img_np = np.array(img)
        # frame = cv2.cvtColor(img_np, cv2.COLOR_BGR2RGB)
        # self.show_img(frame)

        frame = self.get_screen()
        frame_hsv = cv2.cvtColor(frame, cv2.COLOR_BGR2HSV)


        h_min = np.array((0, 0, 253), np.uint8)
        h_max = np.array((255, 0, 255), np.uint8)

        mask = cv2.inRange(frame_hsv, h_min, h_max)
        self.show_img(mask)

        moments = cv2.moments(mask, 1)
        dM01 = moments['m01']
        dM10 = moments['m10']
        dArea = moments['m00']

        print(dM01,dM10,dArea)
        float_x = 0
        float_y = 0

        if dArea > 0:
            float_x = int(dM10 / dArea)
            float_y = int(dM01 / dArea)
            print(float_x,float_y)
            cv2.rectangle(frame, (float_x - 15, float_y - 15), (float_x + 15, float_y + 15), (255, 255, 0), 3)
            # self.show_img(frame)

        return float_x, float_y

    def get_float(self):
        # 加载原始的rgb图像
        img_rgb = self.get_screen()
        # 创建一个原始图像的灰度版本，所有操作在灰度版本中处理，然后在RGB图像中使用相同坐标还原
        img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
        self.show_img(img_gray)
        # 加载将要搜索的图像模板
        template = cv2.imread('f.png', cv2.IMREAD_GRAYSCALE)
        # template_gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
        # height, width = template.shape[:2]
        # size = (int(width * 0.5), int(height * 0.5))
        # template = cv2.resize(template, size, interpolation=cv2.INTER_AREA)
        self.show_img(template)
        # 记录图像模板的尺寸
        w, h = template.shape[::-1]

        res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF)
        # res = cv2.matchTemplate(img_gray, template, cv2.TM_CCOEFF)
        #
        # 'cv2.TM_CCOEFF', 'cv2.TM_CCOEFF_NORMED', 'cv2.TM_CCORR',
        # 'cv2.TM_CCORR_NORMED', 'cv2.TM_SQDIFF', 'cv2.TM_SQDIFF_NORMED'
        # cv2.TM_SQDIFF, cv2.TM_SQDIFF_NORMED 是最小值
        self.show_img(res)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        print(min_val, max_val, min_loc, max_loc)
        print('找到的坐标')

        coordinate = (max_loc[0] + 10, max_loc[1] + 10)  # 左上角的位置
        # top_left = max_loc  # 左上角的位置

        bottom_right = (max_loc[0] + w, max_loc[1] + h)  # 右下角的位置
        cv2.rectangle(img_rgb, max_loc, bottom_right, (255, 255, 0), 3)
        self.show_img(img_rgb)
        return coordinate


    def get_screen(self):
        self.left, self.top, self.right, self.bot = win32gui.GetWindowRect(self.hwnd)
        self.width = self.right - self.left
        self.height = self.bot - self.top
        print(self.left, self.top, self.right, self.bot, self.width, self.height)
        img = pyautogui.screenshot(region=[self.left, self.top, self.width, self.height])  # 分别代表：左上角坐标，宽高
        # 对获取的图片转换成二维矩阵形式，后再将RGB转成BGR
        # 因为imshow,默认通道顺序是BGR，而pyautogui默认是RGB所以要转换一下，不然会有点问题
        img = cv2.cvtColor(np.asarray(img), cv2.COLOR_RGB2BGR)
        # cv2.imwrite("img.jpg", img, [int(cv2.IMWRITE_JPEG_QUALITY), 100])
        # self.show_img(img)
        return img

    def capture(self):
        hwndDC = win32gui.GetWindowDC(self.hwnd)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, self.width, self.height)
        saveDC.SelectObject(saveBitMap)
        saveDC.BitBlt((0, 0), (self.width, self.height), mfcDC, (0, 0), win32con.SRCCOPY)
        signedIntsArray = saveBitMap.GetBitmapBits(True)
        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(self.hwnd, hwndDC)
        im_opencv = np.frombuffer(signedIntsArray, dtype='uint8')
        im_opencv.shape = (self.height, self.width, 4)
        cv2.cvtColor(im_opencv, cv2.COLOR_BGRA2RGB)
        cv2.imwrite("im_opencv.jpg", im_opencv, [int(cv2.IMWRITE_JPEG_QUALITY), 100])  # 保存
        self.show_img(im_opencv)
        return im_opencv

    def show_img(self, img):
        cv2.namedWindow('test')  # 命名窗口
        cv2.imshow("test", img)  # 显示
        cv2.waitKey(0)
        cv2.destroyAllWindows()
        return


if __name__ == "__main__":
    main_window = tk.Tk()  # 实例化出一个父窗口
    app = APP(main_window)
    app.name = '魔兽世界'
    app.name = 'WPS Office'
    # 设置根窗口默认属性
    app.set_init_window()
    main_window.protocol("WM_DELETE_WINDOW", app.on_closing)
    main_window.mainloop()
