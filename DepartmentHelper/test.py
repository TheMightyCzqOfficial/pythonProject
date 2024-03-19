import win32con
import win32gui


def callback(hwnd, hwnd_list):
    hwnd_list.append(hwnd)


def showWindow(windowsName=""):
    hwnd_list = []
    win32gui.EnumWindows(callback, hwnd_list)
    for hwnd in hwnd_list:
        if win32gui.GetWindowText(hwnd) != "":
            print(win32gui.GetWindowText(hwnd))
        if win32gui.GetWindowText(hwnd) == 'exe.win-amd64-3.9':
            win32gui.SetForegroundWindow(hwnd)


if __name__ == '__main__':
    showWindow()
