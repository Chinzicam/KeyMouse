import pyautogui  # 导入 pyautogui 库，用于模拟鼠标和键盘操作
import time  # 导入 time 库，用于控制时间延迟
import openpyxl  # 导入 openpyxl 库，用于读取和写入 Excel 文件
import pyperclip  # 导入 pyperclip 库，用于操作剪贴板
import os  # 导入 os 库，用于文件和操作系统的交互
import configparser  # 导入 configparser 库，用于读取配置文件

# 创建默认的 keyMouse.ini 配置文件
def create_default_ini():
    config = configparser.ConfigParser()
    config['Settings'] = {
        'timeDelay': '1',
        'executionMode': '1',
        'loopCount': '2'
    }
    with open('keyMouse.ini', 'w', encoding='utf-8') as configfile:
        config.write(configfile)

# 创建默认的 orange.xlsx 文件
def create_default_excel():
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.append(['欢迎使用橙子草的Python自动化脚本~（https://github.com/Chinzicam/KeyMouse）'])
    sheet.append(['指令类型（1 单击  2 双击  3 右键  4 输入  5 等待  6滚轮 7 判断 8 键盘键入）',
                  '内容（图片名称.png、输入内容、等待时长/秒）', '重复次数(-1代表一直重复)'])
    wb.save('orange.xlsx')

# 检查并创建必要的文件和文件夹
if not os.path.exists('keyMouse.ini'):
    create_default_ini()

if not os.path.exists('orange.xlsx'):
    create_default_excel()

if not os.path.exists('img'):
    os.makedirs('img')

# 读取配置文件
config = configparser.ConfigParser()
config.read('keyMouse.ini', encoding='utf-8')

# 获取配置项
timeDelay = config.getfloat('Settings', 'timeDelay')
executionMode = config.getint('Settings', 'executionMode')

# 如果是循环模式，读取循环次数
if executionMode == 2:
    loopCount = config.getint('Settings', 'loopCount')
else:
    loopCount = 1

# 定义鼠标事件函数
def mouseClick(clickTimes, lOrR, img, reTry):
    # 检查图像文件是否存在
    if not os.path.exists(img):
        print(f"文件 {img} 不存在，跳过该步骤")
        return
    
    # 根据 reTry 参数的值决定不同的重试逻辑
    if reTry == 1:
        while True:
            # 尝试找到屏幕上的图像位置
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                # 在找到的位置进行鼠标点击操作
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

# 数据检查函数，确保 Excel 中的数据有效
def dataCheck(sheet):
    checkCmd = True
    if sheet.max_row < 3:
        print("没数据啊哥")
        checkCmd = False
    for i in range(3, sheet.max_row + 1):
        cmdType = sheet.cell(row=i, column=1).value
        if cmdType not in [1, 2, 3, 4, 5, 6, 7, 8]:  # 确保操作类型在允许的范围内
            print(f'第{i}行,第1列数据有毛病')
            checkCmd = False
        cmdValue = sheet.cell(row=i, column=2).value
        if cmdType in [1, 2, 3, 7, 8] and not isinstance(cmdValue, str):  # 确保特定类型的值是字符串
            print(f'第{i}行,第2列数据有毛病')
            checkCmd = False
        if cmdType == 4 and not cmdValue:  # 确保类型4的值非空
            print(f'第{i}行,第2列数据有毛病')
            checkCmd = False
        if cmdType in [5, 6] and not isinstance(cmdValue, (int, float)):  # 确保特定类型的值是数字
            print(f'第{i}行,第2列数据有毛病')
            checkCmd = False
    return checkCmd

# 主函数，执行从 Excel 中读取的任务
def mainWork(sheet):
    for i in range(3, sheet.max_row + 1):
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
        elif cmdType == 8:
            keys = sheet.cell(row=i, column=2).value.split('+')
            pyautogui.hotkey(*keys)
            print("按键组合:", '+'.join(keys))
        time.sleep(timeDelay)  # 添加默认等待时间

if __name__ == '__main__':
    file = 'orange.xlsx'  # 定义 Excel 文件名
    wb = openpyxl.load_workbook(filename=file)  # 加载 Excel 工作簿
    sheet = wb.active  # 获取活动工作表
    print('------------------------------~')
    print('欢迎使用橙子草的Python自动化脚本~')
    print('------------------------------~')
    checkCmd = dataCheck(sheet)  # 检查数据有效性
    if checkCmd:
        if executionMode == 1:
            mainWork(sheet)  # 执行一次
        elif executionMode == 2:
            for _ in range(loopCount):
                mainWork(sheet)  # 循环执行
                time.sleep(0.1)
                print("等待0.1秒")
    else:
        print('输入有误或者已经退出!')
