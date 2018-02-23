import win32gui, win32con, win32com.client
import re, traceback
from time import sleep
l_hwnd = []


class cWindow:
    def __init__(self):
        self._hwnd = None

    def SetAsForegroundWindow(self):
        # First, make sure all (other) always-on-top windows are hidden.
        self.hide_always_on_top_windows()
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(self._hwnd)

    def Maximize(self):
        win32gui.ShowWindow(self._hwnd, win32con.SW_MAXIMIZE)

    def _window_enum_callback(self, hwnd, regex):
        '''Pass to win32gui.EnumWindows() to check all open windows'''
        if self._hwnd is None and re.match(
                regex, str(win32gui.GetWindowText(hwnd))) is not None:
            if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(
                    hwnd) and win32gui.IsWindowVisible(hwnd):
                l_hwnd.append(hwnd)

    def find_window_regex(self, regex):
        self._hwnd = None
        win32gui.EnumWindows(self._window_enum_callback, regex)

    def hide_always_on_top_windows(self):
        win32gui.EnumWindows(self._window_enum_callback_hide, None)

    def _window_enum_callback_hide(self, hwnd, unused):
        if hwnd != self._hwnd:  # ignore self
            # Is the window visible and marked as an always-on-top (topmost) window?
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowLong(
                    hwnd, win32con.GWL_EXSTYLE) & win32con.WS_EX_TOPMOST:
                # Ignore windows of class 'Button' (the Start button overlay) and
                # 'Shell_TrayWnd' (the Task Bar).
                className = win32gui.GetClassName(hwnd)
                if not (className == 'Button' or className == 'Shell_TrayWnd'):
                    # Force-minimize the window.
                    # Fortunately, this seems to work even with windows that
                    # have no Minimize button.
                    # Note that if we tried to hide the window with SW_HIDE,
                    # it would disappear from the Task Bar as well.
                    win32gui.ShowWindow(hwnd, win32con.SW_FORCEMINIMIZE)


try:
    regex = ".*"
    print("间隔几秒换下一个窗口？输入数字：")
    tt=input()
    cW = cWindow()
    cW.find_window_regex(regex)
    while True:
        for one_hwnd in l_hwnd:
            cW._hwnd = one_hwnd
            cW.Maximize()
            cW.SetAsForegroundWindow()
            sleep(int(tt))
except:
    f = open("log.txt", "w")
    f.write(traceback.format_exc())
    print(traceback.format_exc())
