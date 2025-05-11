import sys
import win32gui
import win32con
import win32api
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QMessageBox, QSystemTrayIcon, QMenu, QStyle
from PyQt6.QtCore import Qt, QTimer, QRect
from PyQt6.QtGui import QScreen, QIcon, QAction
import logging
import os
import traceback
import ctypes
from ctypes import wintypes, c_void_p
import win32process
from PyQt6.QtWidgets import QScrollArea, QFrame
import uuid
import hashlib
from PyQt6.QtWidgets import QLineEdit, QInputDialog

# 设置日志到控制台和文件
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('window_manager.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

class WindowManager(QMainWindow):
    def __init__(self):
        try:
            super().__init__()
            # 先初始化授权相关属性
            self.machine_code = self.get_machine_code()
            self.license_key = self.generate_license_key(self.machine_code)
            self.is_authorized = self.check_local_license()
            self.target_windows = {}
            self.is_hidden = {}
            self.monitor_timer = QTimer()
            self.monitor_timer.timeout.connect(self.check_window_position)
            self.monitor_timer.start(100)
            self.hook = None
            self.user32 = ctypes.windll.user32
            self.initTray()
            self.setWindowFlag(Qt.WindowType.Tool)
            logging.info('Starting WindowManager initialization')
            self.initUI()
            logging.info('WindowManager initialized successfully')
        except Exception as e:
            logging.error(f'Error in WindowManager initialization: {str(e)}\n{traceback.format_exc()}')
            QMessageBox.critical(None, '错误', f'初始化失败: {str(e)}')
            raise

    def get_machine_code(self):
        return str(uuid.getnode())

    def generate_license_key(self, machine_code):
        h = hashlib.sha256(machine_code.encode('utf-8')).hexdigest()
        return h[-8:].upper()

    def check_local_license(self):
        try:
            license_path = os.path.join(os.getcwd(), '.license.key')
            if os.path.exists(license_path):
                with open(license_path, 'r') as f:
                    key = f.read().strip()
                    return key == self.license_key
            return False
        except Exception as e:
            logging.error(f'Error reading license: {str(e)}')
            return False
    
    def save_local_license(self):
        try:
            license_path = os.path.join(os.getcwd(), '.license.key')
            with open(license_path, 'w') as f:
                f.write(self.license_key)
            # 设置为隐藏文件（仅Windows有效）
            try:
                import ctypes
                FILE_ATTRIBUTE_HIDDEN = 0x02
                ctypes.windll.kernel32.SetFileAttributesW(license_path, FILE_ATTRIBUTE_HIDDEN)
            except Exception as e:
                logging.warning(f'无法设置隐藏属性: {str(e)}')
        except Exception as e:
            logging.error(f'Error saving license: {str(e)}')

    def show_license_dialog(self):
        text, ok = QInputDialog.getText(self, "授权码验证",
            f"本机机器码：{self.machine_code}\n请输入授权码：")
        if ok:
            if text.strip().upper() == self.license_key:
                self.is_authorized = True
                self.save_local_license()
                QMessageBox.information(self, "授权成功", "授权码正确，功能已解锁！")
            else:
                QMessageBox.critical(self, "授权失败", "授权码错误，请联系作者获取授权码。")
        else:
            QMessageBox.warning(self, "未授权", "未输入授权码，部分功能不可用。")

    def initTray(self):
        # 创建系统托盘图标
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon))
        
        # 创建托盘菜单
        tray_menu = QMenu()
        show_action = QAction("显示", self)
        quit_action = QAction("退出", self)
        show_action.triggered.connect(self.show)
        quit_action.triggered.connect(self.realQuit)
        
        tray_menu.addAction(show_action)
        tray_menu.addAction(quit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()
        
        # 设置托盘图标双击事件
        self.tray_icon.activated.connect(self.trayIconActivated)

    def trayIconActivated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.show()
            self.activateWindow()

    def realQuit(self):
        # 真正的退出程序
        self.tray_icon.hide()
        QApplication.quit()

    def closeEvent(self, event):
        # 重写关闭事件，改为最小化到托盘
        if self.hook:
            self.user32.UnhookWindowsHookEx(self.hook)
            self.hook = None
        event.ignore()  # 忽略关闭事件
        self.hide()  # 隐藏窗口

    def clear_windows(self):
        self.target_windows.clear()
        self.is_hidden.clear()
        self.update_window_list()

    def update_window_list(self):
        # 清除现有的窗口列表布局
        if hasattr(self, 'windows_layout'):
            # 清除旧的布局中的所有部件
            while self.windows_layout.count():
                item = self.windows_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
                elif item.layout():
                    # 清理子布局中的部件
                    while item.layout().count():
                        sub_item = item.layout().takeAt(0)
                        if sub_item.widget():
                            sub_item.widget().deleteLater()
            # 从主布局中移除窗口列表布局
            self.main_layout.removeItem(self.windows_layout)
            
        # 创建新的窗口列表布局
        self.windows_layout = QVBoxLayout()
        self.windows_layout.setSpacing(5)  # 设置垂直间距
        
        # 创建一个滚动区域来容纳窗口列表
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        scroll_layout.setSpacing(5)  # 设置垂直间距
        
        for hwnd, title in self.target_windows.items():
            # 为每个窗口创建一个容器部件
            window_widget = QWidget()
            window_layout = QHBoxLayout(window_widget)
            window_layout.setContentsMargins(5, 2, 5, 2)  # 设置边距
            
            # 添加窗口标题和状态
            status = "已隐藏" if self.is_hidden.get(hwnd, False) else "显示中"
            window_label = QLabel(f"• {title} ({status})")
            window_label.setWordWrap(True)  # 允许文本换行
            window_layout.addWidget(window_label, stretch=1)  # 设置拉伸因子
            
            # 添加删除按钮
            delete_btn = QPushButton("删除")
            delete_btn.setMaximumWidth(60)
            delete_btn.clicked.connect(lambda checked, h=hwnd: self.remove_window(h))
            window_layout.addWidget(delete_btn)
            
            # 将窗口部件添加到滚动区域的布局中
            scroll_layout.addWidget(window_widget)
        
        # 如果没有窗口，显示提示文本
        if not self.target_windows:
            empty_label = QLabel("未选择任何窗口")
            empty_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            scroll_layout.addWidget(empty_label)
        
        # 添加弹性空间
        scroll_layout.addStretch()
        
        # 创建滚动区域
        scroll_area = QScrollArea()
        scroll_area.setWidget(scroll_widget)
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.Shape.NoFrame)  # 移除边框
        
        # 将滚动区域添加到主窗口列表布局中
        self.windows_layout.addWidget(scroll_area)
        
        # 将窗口列表布局添加到主布局中
        self.main_layout.insertLayout(1, self.windows_layout)

    def remove_window(self, hwnd):
        """删除单个窗口"""
        if hwnd in self.target_windows:
            del self.target_windows[hwnd]
            if hwnd in self.is_hidden:
                del self.is_hidden[hwnd]
            self.update_window_list()

    def initUI(self):
        self.setWindowTitle('类QQ窗口隐藏器')
        self.setGeometry(100, 100, 500, 300)
        
        # 创建中央窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.main_layout = QVBoxLayout(central_widget)
        
        # 格式化后的程序使用说明，调整授权说明样式
        instruction_text = (
            "<b>用法：</b><br>"
            "1. 点击<b>选择窗口</b>，选择需要隐藏的窗口<br>"
            "2. 把需要隐藏的窗口移动到屏幕右边或顶部边缘，移开鼠标，窗口会自动隐藏<br>"
            "3. 把鼠标移动到屏幕边缘，窗口就会自动弹出来显示，用完后又会自动隐藏<br><br>"
            "<b>授权说明：</b><br>"
            "1. 本软件一机一码，本机机器码：<span style='color:blue;'><b>{machine_code}</b></span><br>"
            "2. 首次运行选择窗口会提示输入授权码<br>"
            "3. 加微信：<span style='color:green;'><b>15050999</b></span> 发机器码，前100名用户免费获取<br>"
            "4. 授权费用：19.9元终身使用"
        ).format(machine_code=self.machine_code)
        self.instruction_label = QLabel(instruction_text)
        self.instruction_label.setWordWrap(True)
        self.instruction_label.setStyleSheet("color: #555; font-size: 13px;")
        self.instruction_label.setTextFormat(Qt.TextFormat.RichText)
        self.main_layout.addWidget(self.instruction_label)
        # 显示机器码
        # self.machine_code_label = QLabel(f"本机机器码：{self.machine_code}")
        # self.main_layout.addWidget(self.machine_code_label)
        self.status_label = QLabel('已选择的窗口:')
        self.main_layout.addWidget(self.status_label)
        button_layout = QHBoxLayout()
        select_btn = QPushButton('选择窗口')
        select_btn.clicked.connect(self.start_window_selection)
        button_layout.addWidget(select_btn)
        clear_btn = QPushButton('清除所有')
        clear_btn.clicked.connect(self.clear_windows)
        button_layout.addWidget(clear_btn)
        self.main_layout.addLayout(button_layout)

    def start_window_selection(self):
        if not self.is_authorized:
            self.show_license_dialog()
            if not self.is_authorized:
                return
        self.status_label.setText('请点击要管理的窗口...')
        
        def mouse_callback(nCode, wParam, lParam):
            if wParam == win32con.WM_LBUTTONDOWN:
                try:
                    cursor_pos = win32gui.GetCursorPos()
                    hwnd = win32gui.WindowFromPoint(cursor_pos)
                    window_text = win32gui.GetWindowText(hwnd)
                    
                    if hwnd not in self.target_windows:
                        self.target_windows[hwnd] = window_text
                        self.is_hidden[hwnd] = False
                        self.update_window_list()
                    
                    # 取消钩子
                    if self.hook:
                        self.user32.UnhookWindowsHookEx(self.hook)
                        self.hook = None
                    
                except Exception as e:
                    logging.error(f'Error in mouse callback: {str(e)}')
                
            return self.user32.CallNextHookEx(None, nCode, wParam, lParam)
        
        # 创建钩子回调函数的C类型
        CMPFUNC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.POINTER(ctypes.c_void_p))
        self.callback_pointer = CMPFUNC(mouse_callback)  # 保存引用防止被垃圾回收
        
        # 设置鼠标钩子
        module_handle = ctypes.c_void_p(win32api.GetModuleHandle(None))
        self.hook = self.user32.SetWindowsHookExA(
            win32con.WH_MOUSE_LL,
            self.callback_pointer,
            module_handle,
            0
        )
        
        if not self.hook:
            logging.error('Failed to set mouse hook')
            QMessageBox.critical(self, '错误', '无法设置鼠标钩子')

    def check_window_position(self):
        if not self.target_windows:
            return
            
        try:
            for hwnd in list(self.target_windows.keys()):
                if not win32gui.IsWindow(hwnd):  # 检查窗口是否还存在
                    del self.target_windows[hwnd]
                    del self.is_hidden[hwnd]
                    self.update_window_list()
                    continue
                    
                rect = win32gui.GetWindowRect(hwnd)
                screen = QApplication.primaryScreen()
                screen_geometry = screen.geometry()
                cursor_pos = win32gui.GetCursorPos()
                
                # 定义边缘检测的灵敏度（像素）
                EDGE_SENSITIVITY = 5
                SHOW_TRIGGER_WIDTH = 5
                
                # 检查窗口是否触及屏幕边缘
                window_right_edge = rect[2]
                window_top_edge = rect[1]
                screen_right_edge = screen_geometry.width()
                
                # 检查鼠标是否在窗口的相应范围内
                is_mouse_in_vertical_range = rect[1] <= cursor_pos[1] <= rect[3]
                is_mouse_in_horizontal_range = rect[0] <= cursor_pos[0] <= rect[2]
                
                # 检查鼠标是否在触发区域
                is_mouse_in_right_trigger = (
                    cursor_pos[0] >= screen_right_edge - SHOW_TRIGGER_WIDTH and
                    is_mouse_in_vertical_range
                )
                is_mouse_in_top_trigger = (
                    cursor_pos[1] <= SHOW_TRIGGER_WIDTH and
                    is_mouse_in_horizontal_range
                )
                
                # 检查鼠标是否在窗口区域内
                is_mouse_in_window = (
                    rect[0] <= cursor_pos[0] <= rect[2] and
                    rect[1] <= cursor_pos[1] <= rect[3]
                )
                
                # 窗口显示/隐藏逻辑
                if not self.is_hidden.get(hwnd, False):
                    if window_right_edge >= screen_right_edge - EDGE_SENSITIVITY:
                        if not is_mouse_in_window:
                            self.hide_window(hwnd, 'right')
                    elif window_top_edge <= EDGE_SENSITIVITY:
                        if not is_mouse_in_window:
                            self.hide_window(hwnd, 'top')
                else:
                    # 如果鼠标在触发区域，无论前台窗口状态如何都显示窗口
                    if (is_mouse_in_right_trigger or is_mouse_in_top_trigger):
                        try:
                            # 强制激活窗口
                            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                            
                            # 使用 keybd_event 模拟按键来强制激活窗口
                            win32api.keybd_event(0, 0, 0, 0)
                            win32api.keybd_event(0, 0, win32con.KEYEVENTF_KEYUP, 0)
                            
                            # 强制设置窗口位置和状态
                            win32gui.SetWindowPos(
                                hwnd,
                                win32con.HWND_TOPMOST,
                                rect[0], rect[1],
                                rect[2] - rect[0], rect[3] - rect[1],
                                win32con.SWP_SHOWWINDOW | win32con.SWP_NOSIZE | win32con.SWP_NOMOVE | 
                                win32con.SWP_ASYNCWINDOWPOS
                            )
                            
                            # 立即调用show_window更新窗口位置
                            self.show_window(hwnd)
                            
                            # 确保窗口完全显示后再取消置顶
                            def delayed_restore():
                                try:
                                    if win32gui.IsWindow(hwnd):
                                        win32gui.SetWindowPos(
                                            hwnd,
                                            win32con.HWND_NOTOPMOST,
                                            rect[0], rect[1],
                                            rect[2] - rect[0], rect[3] - rect[1],
                                            win32con.SWP_SHOWWINDOW | win32con.SWP_NOSIZE | win32con.SWP_NOMOVE |
                                            win32con.SWP_ASYNCWINDOWPOS
                                        )
                                except Exception as e:
                                    logging.error(f"Error in delayed_restore: {str(e)}")
                            
                            QTimer.singleShot(100, delayed_restore)
                            
                        except Exception as e:
                            logging.error(f"Error activating window: {str(e)}")
                        
        except Exception as e:
            logging.error(f"Error in check_window_position: {str(e)}")
            pass

    def restore_window_state(self, hwnd):
        """恢复窗口的正常状态（非置顶）"""
        try:
            rect = win32gui.GetWindowRect(hwnd)
            win32gui.SetWindowPos(
                hwnd,
                win32con.HWND_NOTOPMOST,
                rect[0], rect[1],
                rect[2] - rect[0], rect[3] - rect[1],
                win32con.SWP_SHOWWINDOW | win32con.SWP_NOSIZE | win32con.SWP_NOMOVE
            )
        except Exception as e:
            logging.error(f"Error in restore_window_state: {str(e)}")

    def hide_window(self, hwnd, edge):
        try:
            self.is_hidden[hwnd] = True
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            screen = QApplication.primaryScreen()
            
            if edge == 'right':
                win32gui.MoveWindow(hwnd, screen.geometry().width() - 5, rect[1], width, height, True)
            elif edge == 'top':
                win32gui.MoveWindow(hwnd, rect[0], -height + 5, width, height, True)
            
            self.update_window_list()
                
        except Exception as e:
            logging.error(f"Error in hide_window: {str(e)}")

    def show_window(self, hwnd):
        try:
            self.is_hidden[hwnd] = False
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            screen = QApplication.primaryScreen()
            
            if rect[0] >= screen.geometry().width() - 10:  # 右边隐藏状态
                win32gui.MoveWindow(hwnd, screen.geometry().width() - width, rect[1], width, height, True)
            elif rect[1] < 0:  # 顶部隐藏状态
                win32gui.MoveWindow(hwnd, rect[0], 0, width, height, True)
            
            self.update_window_list()
            
        except Exception as e:
            logging.error(f"Error in show_window: {str(e)}")

    def closeEvent(self, event):
        # 确保在关闭窗口时取消钩子
        if self.hook:
            self.user32.UnhookWindowsHookEx(self.hook)
            self.hook = None
        super().closeEvent(event)

def main():
    try:
        logging.info('Application starting')
        # 确保只有一个QApplication实例
        if not QApplication.instance():
            app = QApplication(sys.argv)
        else:
            app = QApplication.instance()
            
        window_manager = WindowManager()
        window_manager.show()
        logging.info('Main window shown')
        
        # 确保窗口显示在最前面
        window_manager.raise_()
        window_manager.activateWindow()
        
        return app.exec()
    except Exception as e:
        logging.error(f'Critical error in main: {str(e)}\n{traceback.format_exc()}')
        QMessageBox.critical(None, '错误', f'程序启动失败: {str(e)}')
        return 1

if __name__ == '__main__':
    try:
        logging.info('Script started')
        # 添加命令行参数，用于调试
        if len(sys.argv) == 1:
            sys.argv.append('')  # 添加一个空参数
        sys.exit(main())
    except Exception as e:
        logging.error(f'Fatal error: {str(e)}\n{traceback.format_exc()}')
        QMessageBox.critical(None, '致命错误', f'程序崩溃: {str(e)}')
        sys.exit(1)