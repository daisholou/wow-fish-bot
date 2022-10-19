import pyautogui
import win32gui
import win32ui
import win32con
from threading import Timer
import time
import os
import numpy as np
import tkinter as tk
import tkinter.font as tkFont
import tkinter.messagebox as tkmsgbox
import hashlib
import time
import cv2
from PIL import ImageGrab

LOG_LINE_NUM = 0


class APP:

    def __init__(self, init_window_name):
        self.init_window_name = init_window_name
        self.free_hwnd = []
        self.free_name = []
        self.online_hwnd = []
        self.online_name = []
        self.disk_lock = True

    # 设置窗口
    def set_init_window(self):
        self.init_window_name.title("WoW Online V1.00")  # 窗口名
        width = 600
        height = 600
        screenwidth = self.init_window_name.winfo_screenwidth()
        screenheight = self.init_window_name.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        self.init_window_name.geometry(alignstr)
        self.init_window_name.resizable(width=False, height=False)

        self.free_label = tk.Label(self.init_window_name, text="自由窗口")
        self.free_label.place(x=20, y=10, width=70, height=20)

        self.free_wnd_list = tk.Listbox(self.init_window_name, selectmode=tk.MULTIPLE)
        self.free_wnd_list.place(x=20, y=40, width=560, height=60)
        self.refresh()

        self.up_button = tk.Button(self.init_window_name, text="移除", bg="lightblue", width=10, command=self.remove)
        self.up_button.place(x=400, y=120, width=70, height=25)

        self.down_button = tk.Button(self.init_window_name, text="加入", bg="lightblue", width=10, command=self.join)
        self.down_button.place(x=300, y=120, width=70, height=25)

        self.refresh_button = tk.Button(self.init_window_name, text="刷新", bg="lightblue", width=10,
                                        command=self.refresh)
        self.refresh_button.place(x=200, y=120, width=70, height=25)

        self.show_button = tk.Button(self.init_window_name, text="显示窗口", bg="lightblue", width=10,
                                     command=self.show_window)
        self.show_button.place(x=100, y=120, width=70, height=25)

        self.online_label = tk.Label(self.init_window_name, text="托管窗口")
        self.online_label.place(x=20, y=160, width=70, height=20)

        self.online_wnd_list = tk.Listbox(self.init_window_name, selectmode=tk.MULTIPLE)
        self.online_wnd_list.place(x=20, y=190, width=560, height=60)

        # 日志框
        self.log_label = tk.Label(self.init_window_name, text="日志")
        self.log_label.place(x=20, y=270, width=70, height=20)

        self.log_data_Text = tk.Text(self.init_window_name)
        self.log_data_Text.place(x=20, y=300, width=560, height=240)

    # 删除
    def remove(self):
        wnd = self.online_wnd_list.curselection()
        if wnd:
            for e in wnd:
                hwnd = self.online_hwnd[e]
                print(hwnd)
                self.free_hwnd.append(hwnd)
                self.free_name.append(str(hwnd) + ' ' + self.get_window_name(hwnd))  # 给窗口取个名字
                self.free_wnd_list.insert('end', self.free_name[-1])
                self.write_log_to_Text(self.free_name[-1] + ' stop')

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
                self.online_name.append(str(hwnd) + ' ' + self.get_window_name(hwnd))
                self.online_wnd_list.insert('end', self.online_name[-1])
                self.write_log_to_Text(self.online_name[-1] + ' start online')
                # self.get_screen(hwnd)

            for e in reversed(wnd):
                self.free_hwnd.__delitem__(e)
                self.free_name.__delitem__(e)
                self.free_wnd_list.delete(e)

            print(self.free_hwnd)
            print(self.online_hwnd)

    # 刷新数据
    def refresh(self):
        wnd = self.find_window('魔兽世界')

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
            if hwnd not in self.free_hwnd and hwnd not in self.online_hwnd:
                self.free_hwnd.append(hwnd)
                self.free_name.append(str(hwnd) + ' ' + self.get_window_name(hwnd))
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

    def fishing(self, hwnd):
        bait_time = 0
        while hwnd in self.online_hwnd:
            self.write_log_to_Text('%d start fishing', hwnd)
            is_bait = False
            x, y = self.find_float(hwnd)
            if x + y > 0:
                last_x = x
                last_y = y
                # print("Fish on !")
                new_cast_time = time.time()
                is_block = True
                time.sleep(2)
            else:
                if time.localtime(time.time()) - bait_time > 600:
                    is_bait = False
                if not is_bait:
                    pyautogui.press('2')  # 上饵
                    bait_time = time.localtime(time.time())
                    self.write_log_to_Text('%d add bait', hwnd)
                    is_bait = True

                pyautogui.press('1')  # 下竿
                self.write_log_to_Text('%d Lower fishing rod', hwnd)





    def get_window_name(self, hwnd):
        return win32gui.GetWindowText(hwnd)

    def show_window(self):
        wnd = self.free_wnd_list.curselection()
        if wnd:
            for e in wnd:
                hwnd = self.free_hwnd[e]
                print(hwnd)
                win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
                time.sleep(1)

    def get_window_status(self, hwnd):
        status = '未知'
        if self.find_specify_picture(os.getcwd() + "\\wow_login_in.png"):
            status = '人物界面'
        if self.find_specify_picture(os.getcwd() + "\\wow_disconnect.png") or self.find_specify_picture(
                os.getcwd() + "\\wow_disconnect_too.png"):
            status = '断开连接'
        if self.find_specify_picture(os.getcwd() + "\\wow_reconnect.png"):
            status = '重新连接'

    def find_specify_picture(self, targetpicture):
        character = pyautogui.locateOnScreen(targetpicture, confidence=0.5)  # grayscale=True)
        if (character is not None):
            imcenter = pyautogui.center(character)
            print(imcenter)
            return imcenter
        else:
            print("none posion find")
            return None

    def find_float(self, hwnd):

        img = self.get_screen(hwnd)
        img_np = np.array(img)
        frame = cv2.cvtColor(img_np, cv2.COLOR_BGR2RGB)

        self.show_img(frame)

        frame_hsv = cv2.cvtColor(frame, cv2.COLOR_BGR2HSV)

        self.show_img(frame_hsv)

        h_min = np.array((0, 0, 253), np.uint8)
        h_max = np.array((255, 0, 255), np.uint8)

        mask = cv2.inRange(frame_hsv, h_min, h_max)

        moments = cv2.moments(mask, 1)
        dM01 = moments['m01']
        dM10 = moments['m10']
        dArea = moments['m00']

        float_x = 0
        float_y = 0

        if dArea > 0:
            float_x = int(dM10 / dArea)
            float_y = int(dM01 / dArea)
        return float_x, float_y

    def get_screen(self, hwnd):
        left, top, right, bot = win32gui.GetWindowRect(hwnd)
        width = right - left
        height = bot - top
        im_opencv = np.frombuffer(signedIntsArray, dtype='uint8')
        im_opencv.shape = (height, width, 4)
        cv2.cvtColor(im_opencv, cv2.COLOR_BGRA2RGB)
        cv2.imwrite("im_opencv.jpg", im_opencv, [int(cv2.IMWRITE_JPEG_QUALITY), 100])  # 保存
        self.show_img(im_opencv)

        img = ImageGrab.grab(left, top, right, bot)

        self.show_img(img)

        return img

    def show_img(self,img):
        cv2.namedWindow('示例图片')  # 命名窗口
        cv2.imshow("示例图片", img)  # 显示
        cv2.waitKey(0)
        cv2.destroyAllWindows()


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
            self.init_window_name.destroy()


class Wow:
    pass


if __name__ == "__main__":
    main_window = tk.Tk()  # 实例化出一个父窗口
    app = APP(main_window)
    # 设置根窗口默认属性
    app.set_init_window()
    main_window.protocol("WM_DELETE_WINDOW", app.on_closing)
    main_window.mainloop()
