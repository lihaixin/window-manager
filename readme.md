# 编译说明

## 环境准备

1. 确保已安装 Python 3.x
2. 安装所需依赖包:

```
python -m pip install PyQt6 pywin32 pyinstaller


python -m PyInstaller window_manager.spec


git add .
git commit -m "add github actions"
git push -u origin main

python -m PyInstaller  .\window_manager.spec
.\dist\window_manager.exe

```