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
            logging.info('Starting WindowManager initialization')
            self.initUI()
            self.target_windows = {}  # 使用字典存储窗口句柄和标题
            self.is_hidden = {}  # 记录每个窗口的隐藏状态
            self.monitor_timer = QTimer()
            self.monitor_timer.timeout.connect(self.check_window_position)
            self.monitor_timer.start(100)  # 每100ms检查一次窗口位置
            
            # 初始化鼠标钩子
            self.hook = None
            self.user32 = ctypes.windll.user32
            
            # 初始化系统托盘
            self.initTray()
            
            # 设置窗口标志，使其不在任务栏显示
            self.setWindowFlag(Qt.WindowType.Tool)
            
            logging.info('WindowManager initialized successfully')
        except Exception as e:
            logging.error(f'Error in WindowManager initialization: {str(e)}\n{traceback.format_exc()}')
            QMessageBox.critical(None, '错误', f'初始化失败: {str(e)}')
            raise

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
        self.setWindowTitle('窗口管理器')
        self.setGeometry(100, 100, 400, 300)
        
        # 创建中央窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.main_layout = QVBoxLayout(central_widget)
        
        # 添加说明标签
        self.status_label = QLabel('已选择的窗口:')
        self.main_layout.addWidget(self.status_label)
        
        # 窗口列表布局会在update_window_list中动态创建
        
        # 添加按钮布局
        button_layout = QHBoxLayout()
        
        # 添加选择窗口按钮
        select_btn = QPushButton('选择窗口')
        select_btn.clicked.connect(self.start_window_selection)
        button_layout.addWidget(select_btn)
        
        # 添加清除所有按钮
        clear_btn = QPushButton('清除所有')
        clear_btn.clicked.connect(self.clear_windows)
        button_layout.addWidget(clear_btn)
        
        self.main_layout.addLayout(button_layout)

    def start_window_selection(self):
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