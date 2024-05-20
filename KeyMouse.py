import pyautogui
import time
import openpyxl
import pyperclip
import os  # 确保导入 os 模块

# 定义鼠标事件
def mouseClick(clickTimes, lOrR, img, reTry):
    if not os.path.exists(img):
        print(f"文件 {img} 不存在，跳过该步骤")
        return
    if reTry == 1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                break
            print("未找到匹配图片,0.1秒后重试")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                print("重复")
                i += 1
            time.sleep(0.1)

# 数据检查
def dataCheck(sheet):
    checkCmd = True
    if sheet.max_row < 2:
        print("没数据啊哥")
        checkCmd = False
    for i in range(2, sheet.max_row + 1):
        cmdType = sheet.cell(row=i, column=1).value
        if cmdType not in [1, 2, 3, 4, 5, 6, 7]:
            print(f'第{i}行,第1列数据有毛病')
            checkCmd = False
        cmdValue = sheet.cell(row=i, column=2).value
        if cmdType in [1, 2, 3, 7] and not isinstance(cmdValue, str):
            print(f'第{i}行,第2列数据有毛病')
            checkCmd = False
        if cmdType == 4 and not cmdValue:
            print(f'第{i}行,第2列数据有毛病')
            checkCmd = False
        if cmdType in [5, 6] and not isinstance(cmdValue, (int, float)):
            print(f'第{i}行,第2列数据有毛病')
            checkCmd = False
    return checkCmd

# 执行任务
def mainWork(sheet):
    for i in range(2, sheet.max_row + 1):
        cmdType = sheet.cell(row=i, column=1).value
        if cmdType == 1:
            img = 'img/' + sheet.cell(row=i, column=2).value
            reTry = sheet.cell(row=i, column=3).value or 1
            mouseClick(1, "left", img, int(reTry))
            print("单击左键", img)
        elif cmdType == 2:
            img = 'img/' + sheet.cell(row=i, column=2).value
            reTry = sheet.cell(row=i, column=3).value or 1
            mouseClick(2, "left", img, int(reTry))
            print("双击左键", img)
        elif cmdType == 3:
            img = 'img/' + sheet.cell(row=i, column=2).value
            reTry = sheet.cell(row=i, column=3).value or 1
            mouseClick(1, "right", img, int(reTry))
            print("右键", img)
        elif cmdType == 4:
            inputValue = sheet.cell(row=i, column=2).value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            print("输入:", inputValue)
        elif cmdType == 5:
            waitTime = sheet.cell(row=i, column=2).value
            time.sleep(waitTime)
            print("等待", waitTime, "秒")
        elif cmdType == 6:
            scroll = sheet.cell(row=i, column=2).value
            pyautogui.scroll(int(scroll))
            print("滚轮滑动", int(scroll), "距离")
        elif cmdType == 7:
            img = 'img/' + sheet.cell(row=i, column=2).value
            if not os.path.exists(img):
                print(f"文件 {img} 不存在，跳过该步骤")
                continue
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                print(f"图片 {img} 存在，执行后续操作")
                pyautogui.click(location.x, location.y)
            else:
                print(f"图片 {img} 不存在，跳过此步骤")

if __name__ == '__main__':
    file = 'cmd.xlsx'
    wb = openpyxl.load_workbook(filename=file)
    sheet = wb.active
    print('欢迎使用橙子草的Python自动化脚本~')
    checkCmd = dataCheck(sheet)
    if checkCmd:
        key = input('选择功能: 1.做一次 2.循环到死 \n')
        if key == '1':
            mainWork(sheet)
        elif key == '2':
            while True:
                mainWork(sheet)
                time.sleep(0.1)
                print("等待0.1秒")
    else:
        print('输入有误或者已经退出!')
