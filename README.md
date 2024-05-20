# 欢迎使用橙子草的Python自动化脚本

## 使用方法

### 文件结构：

- img文件夹：用于存放图片
- cmd.xlsx：用于配置操作脚本
- KeyMouse.py 或者 KeyMouse.exe：主文件

### 如何运行:双击启动/运行以下代码

```py
python KeyMouse.py
```

## 其他

以下是一些功能8 常用键的名称，更多的键名称可以参考 `pyautogui` 的官方文档：

- 字母键：`'a'`, `'b'`, `'c'`, ... , `'z'`
- 数字键：`'0'`, `'1'`, `'2'`, ... , `'9'`
- 功能键：`'f1'`, `'f2'`, ... , `'f12'`
- 箭头键：`'left'`, `'right'`, `'up'`, `'down'`
- 控制键：`'ctrl'`, `'alt'`, `'shift'`, `'win'`, `'cmd'` (Mac OS)
- 其他键：`'enter'`, `'space'`, `'backspace'`, `'tab'`, `'esc'`, `'delete'`, `'home'`, `'end'`, `'pageup'`, `'pagedown'`

## 打包脚本

### 使用 PyInstaller 打包脚本

```py
pyinstaller --onefile --windowed KeyMouse.py
```

