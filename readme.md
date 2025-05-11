# 介绍
它可以让你把任意窗口（比如聊天软件、工具条等）像QQ那样，自动隐藏到屏幕的右边或顶部边缘，需要时只要把鼠标移动到屏幕边缘，窗口就会自动弹出来显示，用完后又会自动隐藏，非常方便桌面空间管理和多任务操作。
# 编译说明

## 环境准备

1. 确保已安装 Python 3.x
2. 安装所需依赖包:

```
python -m pip install PyQt6 pywin32 pyinstaller


python -m PyInstaller window_manager.spec

```
## 本地trae调试

```
python -m PyInstaller .\window_manager.spec
.\dist\window_manager.exe
git add .
git commit -m "add github actions"
git push -u origin main

python -m PyInstaller  .\window_manager.spec
.\dist\window_manager.exe

```

# 直接运行

从发现版本下载后直接运行即可、见截图

![](./images/main-0.png)

从选择直接从需要在桌面右边和顶部隐藏的窗口（类似QQ），把显示的窗口隐藏即可

最终实现的效果如下：

![](./images/main-1.gif)
