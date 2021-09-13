import time
import win32gui
import win32ui
import win32con
import win32api
from cnocr import CnOcr
import os


def window_capture(filename, pofw, pofh, wpct, hpct, imgfmt):
    """capture specified window and specified screen area's filename.jpg image by win32gui"""
    # 获取指定名称进程的窗口号
    hwnd = win32gui.FindWindow(None, "Rainbow Six")
    # hwnd = 0  # 窗口的编号，0号表示当前活跃窗口
    # 根据窗口句柄获取窗口的设备上下文DC（Device Context）
    hwndDC = win32gui.GetWindowDC(hwnd)
    # 根据窗口的DC获取mfcDC
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
    # mfcDC创建可兼容的DC
    saveDC = mfcDC.CreateCompatibleDC()
    # 创建bigmap准备保存图片
    saveBitMap = win32ui.CreateBitmap()
    # 获取监控器信息
    MoniterDev = win32api.EnumDisplayMonitors(None, None)
    w = MoniterDev[0][2][2]
    h = MoniterDev[0][2][3]
    # print(w, h)  # 图片大小
    # 为bitmap开辟空间

    sspw = int(round(pofw*w))
    ssph = int(round(pofh*h))
    ssw = int(round(wpct*w))
    ssh = int(round(hpct*h))
    print(ssw, ssh)
    saveBitMap.CreateCompatibleBitmap(mfcDC, ssw, ssh)
    # 高度saveDC，将截图保存到saveBitmap中
    saveDC.SelectObject(saveBitMap)
    fullfilename = filename+imgfmt
    # 截取从左上角（0，0）长宽为（w，h）的图片
    saveBitMap.SaveBitmapFile(saveDC, fullfilename)
    saveDC.BitBlt((0, 0), (ssw, ssh), mfcDC,
                  (sspw, ssph), win32con.SRCCOPY)
    saveBitMap.SaveBitmapFile(saveDC, fullfilename)
    """win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)
    """


def main():
    listofimgfpt = ['.jpg', '.png', '.bmp']
    for i in range(7200):
        beg = time.time()
        print(i)
        targetname = "r6ss" + str(i)
        window_capture(targetname, pofw=0.6546875, pofh=0.8791666,
                       wpct=0.10546875, hpct=0.0597222, imgfmt=listofimgfpt[0])
        fulltargetname = targetname + listofimgfpt[0]
        res = CnOcr(model_name='densenet-s-gru', context='cpu',
                    root="C:/Users/Noone/AppData/Roaming/cnocr").ocr(fulltargetname)
        with open("test.txt", "a") as f:
            f.write(str(res)+" \r")
        print(res)
        targetpath = "D:/Onedrive/R6S_spider/"+fulltargetname
        time.sleep(2)
        if res == []:
            os.remove(targetpath)
        end = time.time()
        print(end - beg)

    return 0


if __name__ == '__main__':
    main()
