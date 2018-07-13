import time
from pywinauto.findwindows import find_window
from pywinauto.findbestmatch import find_best_match
from pywinauto.win32functions import SetForegroundWindow
from pywinauto.win32functions import MoveWindow
from pywinauto.application import Application
from pywinauto import mouse
from pywinauto import keyboard


def run():
    app = Application().start("Zoom.exe")
    time.sleep(2)
    zoom_window = find_window(title="Zoom - 免费账号")
    MoveWindow(zoom_window, 0, 0, 400, 400)
    SetForegroundWindow(zoom_window)

    movie_window = find_window(title="电影和电视")
    SetForegroundWindow(movie_window)

    time.sleep(1)
    SetForegroundWindow(zoom_window)
    mouse.click(coords=(150, 150))
    # 等待连接
    time.sleep(10)
    # 开始分享
    keyboard.SendKeys('%s')
    # 选择分享哪个程序
    time.sleep(1)
    keyboard.SendKeys('{TAB}')
    time.sleep(1)
    keyboard.SendKeys('{TAB}')
    time.sleep(1)
    keyboard.SendKeys('{DOWN}')
    time.sleep(1)
    keyboard.SendKeys('{RIGHT}')
    time.sleep(1)
    # 勾选共享电脑声音
    keyboard.SendKeys('{TAB}')
    time.sleep(1)
    keyboard.SendKeys('{SPACE}')
    # 确认分享
    time.sleep(1)
    share_window = find_window(title='选择一个您想共享的窗口或程序')
    SetForegroundWindow(share_window)
    time.sleep(1)
    keyboard.SendKeys('{TAB}')
    time.sleep(1)
    keyboard.SendKeys('{TAB}')
    time.sleep(1)
    # 定位到了“开始共享”按钮
    keyboard.SendKeys('{SPACE}')

    time.sleep(60 * 30)
    # time.sleep(3)
    # 关闭分享
    # keyboard.SendKeys('%s')
    keyboard.SendKeys('%{F4}')
    time.sleep(1)
    # 确认关闭分享
    keyboard.SendKeys('{SPACE}')
    # 关闭 Zoom
    SetForegroundWindow(zoom_window)
    keyboard.SendKeys('%{F4}')
    time.sleep(1)
    keyboard.SendKeys('{SPACE}')


if __name__ == "__main__":
    while True:
        run()
