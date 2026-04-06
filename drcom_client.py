import ctypes
import json
import os
import re
import socket
import sys
import threading
import time
import winreg

import requests
from PyQt6.QtCore import Qt, pyqtSignal, QObject, QTimer
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLineEdit, QPushButton, QCheckBox, QLabel, QSystemTrayIcon, QMenu, QFrame, QMessageBox)


# --- 资源路径助手 ---
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# --- 配置与路径 ---
APP_DATA_DIR = os.path.join(os.getenv('APPDATA', os.path.expanduser("~")), "DrComHelper")
if not os.path.exists(APP_DATA_DIR): os.makedirs(APP_DATA_DIR)
CONFIG_FILE = os.path.join(APP_DATA_DIR, "config.json")
APP_NAME = "DrComFastLogin"
LOGIN_URL = "http://172.25.251.2/drcom/login"
LOGOUT_URL = "http://172.25.251.2:801/eportal/portal/mac/unbind"
ICON_PATH = resource_path("logo.ico")


class ConfigManager:
    @staticmethod
    def load():
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        return {"user": "", "pwd": "", "suffix": "", "auto_login": True, "auto_start": False}

    @staticmethod
    def save(config):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4)

    @staticmethod
    def set_windows_autostart(enable):
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)
            if enable:
                exe_path = os.path.abspath(sys.executable)
                cmd = f'"{exe_path}" --silent'
                winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, cmd)
            else:
                try:
                    winreg.DeleteValue(key, APP_NAME)
                except FileNotFoundError:
                    pass
            winreg.CloseKey(key)
        except Exception as e:
            print(f"自启设置失败: {e}")


class WorkerSignals(QObject):
    status_updated = pyqtSignal(str, str)
    info_received = pyqtSignal(dict)
    # 【新增】专门用于触发UI弹窗的跨线程信号
    auth_error_alert = pyqtSignal(str)


class DrComClient(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = ConfigManager.load()
        self.signals = WorkerSignals()
        self.is_monitoring = False
        self.suffixes = ["@cmcc", "@telecom", "@unicom", ""]

        self.total_online_seconds = 0
        self.live_timer = QTimer()
        self.live_timer.timeout.connect(self.tick_online_time)

        if os.path.exists(ICON_PATH):
            self.setWindowIcon(QIcon(ICON_PATH))

        self.init_ui()
        self.create_tray()

        # 绑定信号与槽
        self.signals.status_updated.connect(self.update_status_display)
        self.signals.info_received.connect(self.update_dashboard)
        self.signals.auth_error_alert.connect(self.show_auth_error_box)  # 【新增】绑定弹窗逻辑

        if self.config.get("auto_login"):
            self.update_status_display("正在自动重连...", "#0078d4")
            threading.Thread(target=self._auto_login_task, daemon=True).start()
        else:
            self.update_status_display("等待操作", "#333")

    def init_ui(self):
        self.setWindowTitle("Dr.COM 校园网助手")
        self.setFixedSize(360, 580)

        # --- 全局精美样式 (QSS) ---
        self.setStyleSheet("""
            QMainWindow { background-color: #f4f7f9; }
            QLabel { color: #2c3e50; font-family: 'Microsoft YaHei'; }
            QLineEdit { 
                border: 2px solid #eef2f7; 
                border-radius: 10px; 
                padding: 10px 15px; 
                background-color: white; 
                font-size: 13px;
                color: #333;
            }
            QLineEdit:focus { border: 2px solid #0078d4; background-color: #fff; }

            QPushButton#LoginBtn {
                background-color: #0078d4;
                color: white;
                border-radius: 12px;
                font-weight: bold;
                font-size: 15px;
            }
            QPushButton#LoginBtn:hover { background-color: #005a9e; }

            QPushButton#LogoutBtn {
                background-color: white;
                color: #e74c3c;
                border: 1px solid #e74c3c;
                border-radius: 8px;
                font-size: 12px;
            }
            QPushButton#LogoutBtn:hover { background-color: #fdf2f2; }

            QCheckBox { spacing: 5px; font-size: 13px; color: #576574; }
            QCheckBox::indicator { width: 20px; height: 20px; }
        """)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(30, 30, 30, 15)
        layout.setSpacing(18)

        # 头部
        header_label = QLabel("Dr.COM Fast")
        header_label.setStyleSheet("font-size: 26px; font-weight: 800; color: #0078d4;")
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(header_label)

        # 输入区域
        input_container = QVBoxLayout()
        input_container.setSpacing(10)
        self.user_input = QLineEdit(self.config['user'])
        self.user_input.setPlaceholderText("请输入校园网账号")
        self.user_input.setFixedHeight(45)
        self.user_input.textChanged.connect(self.sync_settings)
        input_container.addWidget(self.user_input)

        self.pwd_input = QLineEdit(self.config['pwd'])
        self.pwd_input.setPlaceholderText("请输入登录密码")
        self.pwd_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.pwd_input.setFixedHeight(45)
        self.pwd_input.textChanged.connect(self.sync_settings)
        input_container.addWidget(self.pwd_input)
        layout.addLayout(input_container)

        # --- 对称布局：左一右一 ---
        cb_layout = QHBoxLayout()
        self.auto_login_cb = QCheckBox("断线秒连")
        self.auto_login_cb.setChecked(self.config['auto_login'])
        self.auto_login_cb.toggled.connect(self.on_auto_reconnect_toggled)

        self.auto_start_cb = QCheckBox("开机启动")
        self.auto_start_cb.setChecked(self.config['auto_start'])
        self.auto_start_cb.toggled.connect(self.sync_settings)

        cb_layout.addWidget(self.auto_login_cb)
        cb_layout.addStretch()
        cb_layout.addWidget(self.auto_start_cb)
        layout.addLayout(cb_layout)

        self.login_btn = QPushButton("开启极速连接")
        self.login_btn.setObjectName("LoginBtn")
        self.login_btn.setFixedHeight(50)
        self.login_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.login_btn.clicked.connect(self.handle_manual_login_btn)
        layout.addWidget(self.login_btn)

        # 数据面板 (Dashboard Card)
        dash_card = QFrame()
        dash_card.setStyleSheet("background-color: white; border-radius: 18px; border: 1px solid #eef2f7;")
        dash_layout = QVBoxLayout(dash_card)
        dash_layout.setContentsMargins(15, 15, 15, 15)

        self.status_label = QLabel("状态: 准备就绪")
        self.status_label.setStyleSheet("color: #8395a7; font-size: 13px; border:none;")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        dash_layout.addWidget(self.status_label)

        self.online_time_label = QLabel("00:00:00")
        self.online_time_label.setStyleSheet(
            "font-size: 36px; color: #0078d4; font-weight: bold; border:none; font-family: 'Segoe UI', 'Consolas';")
        self.online_time_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        dash_layout.addWidget(self.online_time_label)

        info_grid = QHBoxLayout()
        self.ip_label = QLabel("内网IP: --")
        self.mac_label = QLabel("本机MAC: --")
        for lbl in [self.ip_label, self.mac_label]:
            lbl.setStyleSheet("color: #b2bec3; font-size: 10px; border:none;")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            info_grid.addWidget(lbl)
        dash_layout.addLayout(info_grid)
        layout.addWidget(dash_card)

        # 底部操作
        self.logout_btn = QPushButton("注销当前登录")
        self.logout_btn.setObjectName("LogoutBtn")
        self.logout_btn.setFixedHeight(32)
        self.logout_btn.clicked.connect(self.handle_logout)
        layout.addStretch()
        layout.addWidget(self.logout_btn)

        # --- 高亮署名栏 ---
        footer_layout = QHBoxLayout()
        footer_layout.addStretch()
        self.author_badge = QLabel(" by: 飞行雪绒 ")
        self.author_badge.setStyleSheet("""
            background-color: #0078d4; 
            color: white; 
            border-radius: 10px; 
            font-size: 11px; 
            font-weight: bold;
            padding: 2px 10px;
        """)
        footer_layout.addWidget(self.author_badge)
        footer_layout.addStretch()
        layout.addLayout(footer_layout)

    # --- 逻辑加固 ---
    def sync_settings(self):
        self.config.update({'user': self.user_input.text(), 'pwd': self.pwd_input.text(),
                            'auto_login': self.auto_login_cb.isChecked(), 'auto_start': self.auto_start_cb.isChecked()})
        ConfigManager.save(self.config)
        ConfigManager.set_windows_autostart(self.config['auto_start'])

    def on_auto_reconnect_toggled(self, checked):
        self.sync_settings()
        if checked and not self.check_is_online():
            self.update_status_display("重连中...", "#0078d4")
            threading.Thread(target=self._auto_login_task, daemon=True).start()

    def format_seconds(self, s):
        h, m, s = s // 3600, (s % 3600) // 60, s % 60
        return f"{h:02d}:{m:02d}:{s:02d}"

    def tick_online_time(self):
        self.total_online_seconds += 1
        self.online_time_label.setText(self.format_seconds(self.total_online_seconds))

    def update_dashboard(self, data):
        self.ip_label.setText(f"内网IP: {data.get('ss5', '--')}")
        self.mac_label.setText(f"本机MAC: {data.get('ss4', '--')}")
        self.total_online_seconds = int(data.get('aolno', 0))
        if not self.live_timer.isActive(): self.live_timer.start(1000)

    def create_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon(ICON_PATH) if os.path.exists(ICON_PATH) else QApplication.style().standardIcon(
            QApplication.style().StandardPixmap.SP_DriveNetIcon))
        menu = QMenu()
        menu.addAction("打开助手").triggered.connect(self.show)
        menu.addAction("安全退出").triggered.connect(QApplication.instance().quit)
        self.tray_icon.setContextMenu(menu)
        self.tray_icon.show()

    def update_status_display(self, text, color):
        self.status_label.setText(f"状态: {text}")
        self.status_label.setStyleSheet(f"color: {color}; font-weight: bold; border:none;")

    # 【新增】处理严重的账号异常弹窗
    def show_auth_error_box(self, error_msg):
        # 取消自动登录，防止陷入错误死循环被拉黑
        self.auto_login_cb.setChecked(False)
        self.sync_settings()
        # 弹出醒目警告
        QMessageBox.critical(self, "认证拦截",
                             f"无法登录校园网，系统拦截异常：\n\n【{error_msg}】\n\n请检查您的账号密码或网络状态。")
        self.show()  # 如果软件在托盘，强行唤醒显示错误

    def check_is_online(self):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(1.2)
                s.connect(("223.5.5.5", 53))
            return True
        except:
            return False

    # 【深度强化】Dr.COM 报错黑话翻译机
    def parse_drcom_jsonp(self, text):
        match = re.search(r'dr1003\((.*)\)', text)
        if match:
            try:
                data = json.loads(match.group(1))
                self.signals.info_received.emit(data)

                # 提取网关原味报错信息
                raw_msg = data.get("msga", "")

                # 成功判断 (注意包含你发现的 clientip online)
                if "online" in raw_msg or data.get("result") == 1:
                    return True, "连接成功"

                # 失败翻译
                friendly_msg = raw_msg if raw_msg else "验证失败"
                if "userid error2" in raw_msg:
                    friendly_msg = "密码错误 (或账号异常)"
                elif "userid error1" in raw_msg:
                    friendly_msg = "账号不存在 (请检查后缀)"
                elif "ldap auth error" in raw_msg:
                    friendly_msg = "学校认证系统宕机"
                elif "overdue" in raw_msg.lower():
                    friendly_msg = "账号已欠费"
                elif "ip_exist_error" in raw_msg or "已在线" in raw_msg:
                    friendly_msg = "账号已在其他设备登录"

                return False, friendly_msg
            except:
                pass
        return False, "解析失败"

    # 【加固】保留最真实的报错
    def auto_carrier_discovery(self):
        success, msg = self.send_login_request()
        if success: return True, msg

        last_error = msg

        # 如果明确是密码错误，不需要再换运营商后缀测试了，直接阻断
        if "密码" in msg or "欠费" in msg or "其他设备" in msg:
            return False, msg

        # 继续尝试后缀
        for sfx in self.suffixes:
            self.signals.status_updated.emit(f"检测线路 {sfx}...", "#f5a623")
            ok, cur_msg = self.send_login_request(sfx)
            if ok:
                self.config['suffix'] = sfx
                ConfigManager.save(self.config)
                return True, f"重连成功 ({sfx})"
            # 记录最新的有效报错
            if cur_msg and "失败" not in cur_msg:
                last_error = cur_msg

        return False, last_error

    def send_login_request(self, suffix=None):
        u, p = self.config['user'], self.config['pwd']
        sfx = suffix if suffix is not None else self.config.get("suffix", "")
        params = {"callback": "dr1003", "DDDDD": f"{u}{sfx}", "upass": p, "0MKKey": "123456", "terminal_type": "1",
                  "lang": "zh-cn"}
        try:
            res = requests.get(LOGIN_URL, params=params, timeout=5)
            return self.parse_drcom_jsonp(res.text)
        except:
            return False, "网关未响应，请检查是否连接校园WiFi"

    def handle_manual_login_btn(self):
        self.sync_settings()
        self.update_status_display("正在连接...", "#0078d4")
        threading.Thread(target=self._auto_login_task, daemon=True).start()

    def handle_logout(self):
        self.live_timer.stop()
        self.online_time_label.setText("00:00:00")
        threading.Thread(target=lambda: requests.get(LOGOUT_URL, params={"callback": "dr1003",
                                                                         "user_account": f"{self.config['user']}{self.config['suffix']}"}),
                         daemon=True).start()
        self.update_status_display("已注销连接", "#666")

    def _auto_login_task(self):
        if not self.check_is_online():
            success, msg = self.auto_carrier_discovery()
            self.signals.status_updated.emit(msg, "#28a745" if success else "#d93025")

            # 【新增】如果是账号/密码等硬性错误，触发弹窗报警
            if not success and ("密码" in msg or "不存在" in msg or "欠费" in msg):
                self.signals.auth_error_alert.emit(msg)
        else:
            self.signals.status_updated.emit("网络已就绪", "#28a745")
            self.send_login_request()

        self.start_monitoring_thread()

    def start_monitoring_thread(self):
        if self.is_monitoring: return
        self.is_monitoring = True

        def _loop():
            while True:
                # 只有在断线秒连勾选时，才执行重连
                if self.auto_login_cb.isChecked() and not self.check_is_online():
                    self.signals.status_updated.emit("断线秒连中...", "#0078d4")
                    success, msg = self.auto_carrier_discovery()
                    if success:
                        self.signals.status_updated.emit("重连成功！", "#28a745")
                    else:
                        self.signals.status_updated.emit(f"重连失败", "#d93025")
                        # 如果自动重连发现密码改了或者欠费了，立刻弹窗并停止自动重连
                        if "密码" in msg or "不存在" in msg or "欠费" in msg:
                            self.signals.auth_error_alert.emit(msg)
                time.sleep(5)

        threading.Thread(target=_loop, daemon=True).start()

    def closeEvent(self, e):
        if self.tray_icon.isVisible(): self.hide(); e.ignore()


if __name__ == "__main__":
    mutex_name = "Global\\DrComHelper_MaJieRu_Unique"
    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
    if ctypes.windll.kernel32.GetLastError() == 183:
        if "--silent" not in sys.argv:
            temp_app = QApplication(sys.argv)
            QMessageBox.information(None, "运行提示", "飞行雪绒定制版助手已在后台为您护航！")
        sys.exit(0)

    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('mjr.drcom.fast.v1')
    except:
        pass

    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    client = DrComClient()
    if "--silent" not in sys.argv: client.show()
    sys.exit(app.exec())
