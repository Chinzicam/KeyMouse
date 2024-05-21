# Python自动化脚本

> 时隔三年，当初在首页看见水哥视频时，编程能力还不是很熟练，前两天又在首页刷到这个视频了==> [5分钟，教你做个自动化软件拿来办公、刷副本、回微信 | 源码公开，开箱即用_哔哩哔哩_bilibili](https://www.bilibili.com/video/BV1T34y1o73U/?spm_id_from=333.999.0.0) 
>
> 现在一来确实有这个自动化脚本的需求，二来当初这个脚本功能确实不是很完善，现在也有能力更改了，因此花了一上午将脚本重新完善，部分代码重构并添加了一些实用功能

## 使用方法

### 文件结构：

- img文件夹：用于存放图片
- orange.xlsx：用于配置操作脚本
- keyMouse.ini：配置文件
- KeyMouse.py 或者 KeyMouse.exe：主文件

### 如何运行:双击启动exe/运行以下代码

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
pyinstaller --onefile KeyMouse.py
```

