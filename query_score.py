import sys
import re
import requests
import time
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from urllib.parse import quote, unquote
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
                             QGroupBox, QLabel, QLineEdit, QPushButton, QTextEdit, QTabWidget,
                             QSpinBox, QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
                             QCheckBox, QProgressBar, QStyleFactory, QPlainTextEdit, QSplitter)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QPalette, QColor, QIcon, QTextCursor

# 主应用类
class ScoreCheckerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("软考成绩自动查询工具")
        self.setGeometry(100, 100, 1000, 800)
        self.setWindowIcon(QIcon(":icon.png"))

        # 设置应用样式
        self.setStyle(QStyleFactory.create("Fusion"))
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(53, 53, 53))
        palette.setColor(QPalette.WindowText, Qt.white)
        palette.setColor(QPalette.Base, QColor(35, 35, 35))
        palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
        palette.setColor(QPalette.ToolTipBase, Qt.white)
        palette.setColor(QPalette.ToolTipText, Qt.white)
        palette.setColor(QPalette.Text, Qt.white)
        palette.setColor(QPalette.Button, QColor(53, 53, 53))
        palette.setColor(QPalette.ButtonText, Qt.white)
        palette.setColor(QPalette.BrightText, Qt.red)
        palette.setColor(QPalette.Highlight, QColor(142, 45, 197).lighter())
        palette.setColor(QPalette.HighlightedText, Qt.black)
        self.setPalette(palette)

        # 初始化UI
        self.init_ui()

        # 初始化默认值
        self.load_default_settings()

    def init_ui(self):
        # 创建主控件和布局
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)

        # 创建分割器
        splitter = QSplitter(Qt.Vertical)
        main_layout.addWidget(splitter)

        # 创建上半部分：设置区域
        settings_widget = QWidget()
        settings_layout = QVBoxLayout(settings_widget)

        # 创建选项卡
        tabs = QTabWidget()
        settings_layout.addWidget(tabs)

        # 基本设置选项卡
        basic_tab = QWidget()
        basic_layout = QVBoxLayout(basic_tab)
        tabs.addTab(basic_tab, "基本设置")

        # cURL导入区域
        curl_group = QGroupBox("导入cURL命令")
        curl_layout = QVBoxLayout(curl_group)

        self.curl_input = QPlainTextEdit()
        self.curl_input.setPlaceholderText("粘贴cURL命令到这里...")
        self.curl_input.setStyleSheet("font-family: Consolas; font-size: 10pt;")
        curl_layout.addWidget(self.curl_input)

        parse_btn_layout = QHBoxLayout()
        self.parse_curl_btn = QPushButton("解析cURL")
        self.parse_curl_btn.setStyleSheet("background-color: #2196F3; color: white;")
        self.parse_curl_btn.clicked.connect(self.parse_curl)
        parse_btn_layout.addWidget(self.parse_curl_btn)

        clear_curl_btn = QPushButton("清空")
        clear_curl_btn.clicked.connect(self.clear_curl)
        parse_btn_layout.addWidget(clear_curl_btn)

        curl_layout.addLayout(parse_btn_layout)
        basic_layout.addWidget(curl_group)

        # 查询设置组
        query_group = QGroupBox("查询设置")
        query_layout = QGridLayout(query_group)

        # URL设置
        query_layout.addWidget(QLabel("查询URL:"), 0, 0)
        self.url_input = QLineEdit()
        query_layout.addWidget(self.url_input, 0, 1, 1, 3)

        # 考试阶段
        query_layout.addWidget(QLabel("考试阶段:"), 1, 0)
        self.stage_input = QLineEdit()
        query_layout.addWidget(self.stage_input, 1, 1, 1, 3)

        # 请求间隔
        query_layout.addWidget(QLabel("请求间隔(秒):"), 2, 0)
        self.interval_input = QSpinBox()
        self.interval_input.setRange(10, 3600)
        self.interval_input.setSingleStep(30)
        query_layout.addWidget(self.interval_input, 2, 1)

        # 最大尝试次数
        query_layout.addWidget(QLabel("最大尝试次数:"), 2, 2)
        self.max_attempts_input = QSpinBox()
        self.max_attempts_input.setRange(0, 10000)
        self.max_attempts_input.setSpecialValueText("无限")
        query_layout.addWidget(self.max_attempts_input, 2, 3)

        # 失败提醒间隔
        query_layout.addWidget(QLabel("失败提醒间隔:"), 3, 0)
        self.fail_interval_input = QSpinBox()
        self.fail_interval_input.setRange(1, 100)
        self.fail_interval_input.setValue(6)
        query_layout.addWidget(self.fail_interval_input, 3, 1)

        # 用户代理
        query_layout.addWidget(QLabel("User-Agent:"), 4, 0)
        self.user_agent_input = QLineEdit()
        query_layout.addWidget(self.user_agent_input, 4, 1, 1, 3)

        basic_layout.addWidget(query_group)

        # Cookie设置组
        cookie_group = QGroupBox("Cookie设置")
        cookie_layout = QVBoxLayout(cookie_group)

        # Cookie表格
        self.cookie_table = QTableWidget()
        self.cookie_table.setColumnCount(2)
        self.cookie_table.setHorizontalHeaderLabels(["键", "值"])
        self.cookie_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.cookie_table.setRowCount(5)
        cookie_layout.addWidget(self.cookie_table)

        # Cookie操作按钮
        cookie_btn_layout = QHBoxLayout()
        add_cookie_btn = QPushButton("添加行")
        add_cookie_btn.clicked.connect(self.add_cookie_row)
        cookie_btn_layout.addWidget(add_cookie_btn)

        clear_cookie_btn = QPushButton("清空")
        clear_cookie_btn.clicked.connect(self.clear_cookie_table)
        cookie_btn_layout.addWidget(clear_cookie_btn)

        cookie_layout.addLayout(cookie_btn_layout)
        basic_layout.addWidget(cookie_group)

        # 邮件设置选项卡
        mail_tab = QWidget()
        mail_layout = QVBoxLayout(mail_tab)
        tabs.addTab(mail_tab, "邮件设置")

        # 邮件服务器设置
        server_group = QGroupBox("邮件服务器设置")
        server_layout = QGridLayout(server_group)

        server_layout.addWidget(QLabel("SMTP服务器:"), 0, 0)
        self.smtp_server_input = QLineEdit()
        server_layout.addWidget(self.smtp_server_input, 0, 1)

        server_layout.addWidget(QLabel("SMTP端口:"), 0, 2)
        self.smtp_port_input = QSpinBox()
        self.smtp_port_input.setRange(1, 65535)
        self.smtp_port_input.setValue(465)
        server_layout.addWidget(self.smtp_port_input, 0, 3)

        server_layout.addWidget(QLabel("发件邮箱:"), 1, 0)
        self.sender_email_input = QLineEdit()
        server_layout.addWidget(self.sender_email_input, 1, 1)

        server_layout.addWidget(QLabel("邮箱授权码:"), 1, 2)
        self.sender_pwd_input = QLineEdit()
        self.sender_pwd_input.setEchoMode(QLineEdit.Password)
        server_layout.addWidget(self.sender_pwd_input, 1, 3)

        server_layout.addWidget(QLabel("收件邮箱:"), 2, 0)
        self.receiver_email_input = QLineEdit()
        server_layout.addWidget(self.receiver_email_input, 2, 1, 1, 3)

        mail_layout.addWidget(server_group)

        # 邮件测试
        test_group = QGroupBox("邮件测试")
        test_layout = QVBoxLayout(test_group)

        self.test_email_btn = QPushButton("测试邮件发送")
        self.test_email_btn.clicked.connect(self.send_test_email)
        test_layout.addWidget(self.test_email_btn)

        self.test_result_label = QLabel("")
        self.test_result_label.setAlignment(Qt.AlignCenter)
        test_layout.addWidget(self.test_result_label)

        mail_layout.addWidget(test_group)
        mail_layout.addStretch()

        # 添加设置区域到分割器
        splitter.addWidget(settings_widget)

        # 创建下半部分：日志输出区域
        log_widget = QWidget()
        log_layout = QVBoxLayout(log_widget)

        log_group = QGroupBox("运行日志")
        log_group_layout = QVBoxLayout(log_group)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setFont(QFont("Consolas", 10))
        self.log_output.setStyleSheet("background-color: #1e1e1e; color: #d4d4d4;")
        log_group_layout.addWidget(self.log_output)

        # 控制按钮区域
        control_layout = QHBoxLayout()

        self.start_btn = QPushButton("开始查询")
        self.start_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        self.start_btn.clicked.connect(self.start_checking)
        control_layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("停止查询")
        self.stop_btn.setStyleSheet("background-color: #F44336; color: white; font-weight: bold;")
        self.stop_btn.clicked.connect(self.stop_checking)
        self.stop_btn.setEnabled(False)
        control_layout.addWidget(self.stop_btn)

        self.clear_log_btn = QPushButton("清空日志")
        self.clear_log_btn.clicked.connect(self.clear_log)
        control_layout.addWidget(self.clear_log_btn)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #444;
                border-radius: 3px;
                background-color: #2d2d2d;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                width: 10px;
            }
        """)

        # 状态栏
        self.status_label = QLabel("就绪")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("font-weight: bold; color: #64B5F6;")

        log_layout.addWidget(log_group)
        log_layout.addLayout(control_layout)
        log_layout.addWidget(self.progress_bar)
        log_layout.addWidget(self.status_label)

        # 添加日志区域到分割器
        splitter.addWidget(log_widget)

        # 设置分割比例
        splitter.setSizes([400, 400])

        # 设置中心控件
        self.setCentralWidget(central_widget)

        # 初始化工作线程
        self.worker = None

    def load_default_settings(self):
        """加载默认设置"""
        self.url_input.setText("https://bm.ruankao.org.cn/query/score/result")
        self.stage_input.setText("2025年上半年")
        self.interval_input.setValue(600)
        self.max_attempts_input.setValue(0)
        self.user_agent_input.setText("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/137.0.0.0 Safari/537.36")

        # 设置默认Cookie
        cookies = []

        self.cookie_table.setRowCount(len(cookies))
        for i, (key, value) in enumerate(cookies):
            self.cookie_table.setItem(i, 0, QTableWidgetItem(key))
            self.cookie_table.setItem(i, 1, QTableWidgetItem(value))

        # 邮件设置
        self.smtp_server_input.setText("")
        self.smtp_port_input.setValue(465)
        self.sender_email_input.setText("")
        self.sender_pwd_input.setText("")
        self.receiver_email_input.setText("")

    def add_cookie_row(self):
        """添加Cookie行"""
        row_count = self.cookie_table.rowCount()
        self.cookie_table.insertRow(row_count)
        self.log_message("添加了一个新的Cookie行")

    def clear_cookie_table(self):
        """清空Cookie表格"""
        self.cookie_table.setRowCount(0)
        self.log_message("已清空Cookie表格")

    def clear_curl(self):
        """清空cURL输入"""
        self.curl_input.clear()
        self.log_message("已清空cURL输入框")

    def clear_log(self):
        """清空日志"""
        self.log_output.clear()
        self.log_message("日志已清空")

    def log_message(self, message):
        """记录日志消息"""
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}"

        # 在界面上显示日志
        self.log_output.append(log_entry)
        self.log_output.moveCursor(QTextCursor.End)

        # 同时在控制台打印
        print(log_entry)

    def get_cookie_string(self):
        """从表格中获取Cookie字符串"""
        cookie_items = []
        for row in range(self.cookie_table.rowCount()):
            key_item = self.cookie_table.item(row, 0)
            value_item = self.cookie_table.item(row, 1)

            if key_item and value_item and key_item.text() and value_item.text():
                cookie_items.append(f"{key_item.text()}={value_item.text()}")

        return "; ".join(cookie_items)

    def parse_curl(self):
        """解析cURL命令"""
        curl_text = self.curl_input.toPlainText().strip()
        if not curl_text:
            self.log_message("错误: cURL输入为空")
            QMessageBox.warning(self, "解析错误", "请粘贴有效的cURL命令")
            return

        try:
            # 解析URL
            url_match = re.search(r"curl\s+['\"](https?://[^'\"]+)['\"]", curl_text)
            if not url_match:
                self.log_message("错误: 无法解析URL")
                raise ValueError("无法解析URL")

            url = url_match.group(1)
            self.url_input.setText(url)
            self.log_message(f"解析URL: {url}")

            # 解析请求头
            headers = {}
            header_matches = re.findall(r"-H\s+['\"]([^'\"]+)['\"]", curl_text)
            for header in header_matches:
                if ':' in header:
                    key, value = header.split(':', 1)
                    key = key.strip()
                    value = value.strip()
                    headers[key] = value
                    self.log_message(f"解析请求头: {key} = {value}")

            # 更新特定请求头
            if "User-Agent" in headers:
                self.user_agent_input.setText(headers["User-Agent"])
                self.log_message(f"设置User-Agent: {headers['User-Agent']}")

            # 解析Cookie
            cookie_match = re.search(r"-b\s+['\"]([^'\"]+)['\"]", curl_text)
            if not cookie_match:
                cookie_match = re.search(r"--cookie\s+['\"]([^'\"]+)['\"]", curl_text)

            if cookie_match:
                cookie_str = cookie_match.group(1)
                self.log_message(f"解析Cookie字符串: {cookie_str}")

                # 清除现有Cookie
                self.cookie_table.setRowCount(0)

                # 解析Cookie键值对
                cookies = []
                for cookie in cookie_str.split(';'):
                    cookie = cookie.strip()
                    if '=' in cookie:
                        key, value = cookie.split('=', 1)
                        cookies.append((key.strip(), value.strip()))

                # 填充Cookie表格
                self.cookie_table.setRowCount(len(cookies))
                for i, (key, value) in enumerate(cookies):
                    self.cookie_table.setItem(i, 0, QTableWidgetItem(key))
                    self.cookie_table.setItem(i, 1, QTableWidgetItem(value))
                    self.log_message(f"添加Cookie: {key} = {value}")

            # 解析请求数据
            data_match = re.search(r"--data-raw\s+['\"]([^'\"]+)['\"]", curl_text)
            if not data_match:
                data_match = re.search(r"--data\s+['\"]([^'\"]+)['\"]", curl_text)

            if data_match:
                data_str = data_match.group(1)
                self.log_message(f"解析请求数据: {data_str}")

                # 尝试解析考试阶段
                if 'stage=' in data_str:
                    stage_match = re.search(r"stage=([^&]+)", data_str)
                    if stage_match:
                        stage_encoded = stage_match.group(1)
                        stage = unquote(stage_encoded)
                        self.stage_input.setText(stage)
                        self.log_message(f"解析考试阶段: {stage}")

            self.log_message("cURL解析完成")
            QMessageBox.information(self, "解析成功", "cURL命令已成功解析并填充到相应字段")

        except Exception as e:
            self.log_message(f"解析cURL时出错: {str(e)}")
            QMessageBox.critical(self, "解析错误", f"解析cURL时出错:\n{str(e)}")

    def send_test_email(self):
        """发送测试邮件"""
        smtp_server = self.smtp_server_input.text()
        smtp_port = self.smtp_port_input.value()
        sender_email = self.sender_email_input.text()
        sender_pwd = self.sender_pwd_input.text()
        receiver_email = self.receiver_email_input.text()

        if not all([smtp_server, smtp_port, sender_email, sender_pwd, receiver_email]):
            self.log_message("邮件测试失败: 请填写完整的邮件配置")
            self.test_result_label.setText("请填写完整的邮件配置")
            self.test_result_label.setStyleSheet("color: red;")
            return

        self.log_message("开始发送测试邮件...")

        try:
            subject = "软考成绩查询工具测试邮件"
            content = """
            <h2>测试邮件</h2>
            <p>这是一封来自软考成绩查询工具的测试邮件。</p>
            <p>如果您能收到此邮件，说明您的邮件配置是正确的。</p>
            <p style="color:gray">发送时间: {}</p>
            """.format(time.strftime('%Y-%m-%d %H:%M:%S'))

            msg = MIMEText(content, "html", "utf-8")
            msg["Subject"] = Header(subject, "utf-8")
            msg["From"] = sender_email
            msg["To"] = receiver_email

            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(sender_email, sender_pwd)
                server.sendmail(sender_email, [receiver_email], msg.as_string())

            self.log_message("测试邮件发送成功！")
            self.test_result_label.setText("测试邮件发送成功！")
            self.test_result_label.setStyleSheet("color: green;")
        except Exception as e:
            error_msg = f"邮件发送失败: {str(e)}"
            self.log_message(error_msg)
            self.test_result_label.setText(error_msg)
            self.test_result_label.setStyleSheet("color: red;")

    def start_checking(self):
        """开始查询"""
        # 验证必要参数
        if not self.url_input.text():
            self.log_message("错误: 请填写查询URL")
            QMessageBox.warning(self, "参数缺失", "请填写查询URL")
            return

        if not self.stage_input.text():
            self.log_message("错误: 请填写考试阶段")
            QMessageBox.warning(self, "参数缺失", "请填写考试阶段")
            return

        if self.cookie_table.rowCount() == 0:
            self.log_message("错误: 请至少添加一个Cookie")
            QMessageBox.warning(self, "参数缺失", "请至少添加一个Cookie")
            return

        # 获取参数
        params = {
            "url": self.url_input.text(),
            "stage": self.stage_input.text(),
            "interval": self.interval_input.value(),
            "max_attempts": self.max_attempts_input.value(),
            "fail_interval": self.fail_interval_input.value(),
            "user_agent": self.user_agent_input.text(),
            "cookie": self.get_cookie_string(),
            "smtp_server": self.smtp_server_input.text(),
            "smtp_port": self.smtp_port_input.value(),
            "sender_email": self.sender_email_input.text(),
            "sender_pwd": self.sender_pwd_input.text(),
            "receiver_email": self.receiver_email_input.text()
        }

        # 检查邮件配置是否完整
        email_fields = [
            params["smtp_server"],
            params["sender_email"],
            params["sender_pwd"],
            params["receiver_email"]
        ]
        enable_email = all(email_fields)  # 所有字段都有值才启用邮件

        if enable_email:
            self.log_message("邮件配置完整，启用邮件通知功能")
        else:
            self.log_message("邮件配置不完整，禁用邮件通知功能")

        # 记录配置
        self.log_message("启动查询任务...")
        self.log_message(f"目标URL: {params['url']}")
        self.log_message(f"考试阶段: {params['stage']}")
        self.log_message(f"请求间隔: {params['interval']}秒")
        self.log_message(f"最大尝试次数: {'无限' if params['max_attempts'] == 0 else params['max_attempts']}")
        self.log_message(f"失败提醒间隔: 每 {params['fail_interval']} 次失败发送提醒")

        # 添加邮件配置状态到参数
        params["enable_email"] = enable_email

        # 启动工作线程
        self.worker = WorkerThread(params)
        self.worker.log_signal.connect(self.log_message)
        self.worker.status_signal.connect(self.update_status)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished.connect(self.worker_finished)

        # 更新UI状态
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.status_label.setText("正在查询...")
        self.progress_bar.setValue(0)

        # 启动线程
        self.worker.start()
        self.log_message(f"开始查询 {params['stage']} 软考成绩...")

    def stop_checking(self):
        """停止查询"""
        if self.worker and self.worker.isRunning():
            self.log_message("正在停止查询任务...")
            self.worker.stop()
            self.status_label.setText("正在停止...")

    def worker_finished(self):
        """工作线程完成"""
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.status_label.setText("已停止")
        self.progress_bar.setValue(0)
        self.worker = None
        self.log_message("查询任务已停止")

    def update_status(self, message):
        """更新状态标签"""
        self.status_label.setText(message)

    def update_progress(self, value):
        """更新进度条"""
        self.progress_bar.setValue(value)


# 工作线程类
class WorkerThread(QThread):
    log_signal = pyqtSignal(str)
    status_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)

    def __init__(self, params):
        super().__init__()
        self.params = params
        self.running = True
        self.attempt_count = 0
        self.fail_count = 0

    def run(self):
        """线程主函数"""
        self.log_signal.emit(f"任务启动，最大尝试次数: {'无限' if self.params['max_attempts'] == 0 else self.params['max_attempts']}")
        self.log_signal.emit(f"邮件通知功能: {'启用' if self.params['enable_email'] else '禁用'}")

        while self.running and (self.params["max_attempts"] == 0 or self.attempt_count < self.params["max_attempts"]):
            self.attempt_count += 1
            self.fail_count += 1

            current_time = time.strftime('%Y-%m-%d %H:%M:%S')
            self.log_signal.emit(f"\n尝试 #{self.attempt_count} [{current_time}]")
            self.status_signal.emit(f"尝试 #{self.attempt_count}")

            # 更新进度
            progress = min(100, self.attempt_count % 100)
            self.progress_signal.emit(progress)

            # 执行查询
            self.log_signal.emit("正在发送查询请求...")
            success, result = self.query_score()

            if success:
                self.log_signal.emit("🎉 成绩已公布！")
                self.log_signal.emit(f"成绩信息：{result}")

                # 发送成功通知邮件（如果邮件功能启用）
                if self.params["enable_email"]:
                    self.log_signal.emit("正在发送成绩通知邮件...")
                    if self.send_email("成绩已公布", result):
                        self.log_signal.emit("邮件通知已发送")
                else:
                    self.log_signal.emit("邮件功能未启用，跳过发送成绩通知")

                self.status_signal.emit("成绩已公布")
                self.progress_signal.emit(100)
                break

            self.log_signal.emit(f"查询结果: {result}")

            # 检查是否需要发送失败提醒（如果邮件功能启用）
            if self.params["enable_email"] and self.fail_count % self.params["fail_interval"] == 0:
                self.log_signal.emit(f"达到失败提醒间隔 ({self.params['fail_interval']}次)，发送提醒邮件")
                subject = f"软考成绩查询失败提醒 - 已尝试{self.attempt_count}次"
                content = f"""
                <h3>软考成绩查询失败提醒</h3>
                <p>考试阶段：{self.params['stage']}</p>
                <p>已尝试查询次数：{self.attempt_count}</p>
                <p>最近一次错误：{result}</p>
                <p>最后尝试时间：{current_time}</p>
                """
                self.send_email(subject, content)
            elif self.fail_count % self.params["fail_interval"] == 0:
                self.log_signal.emit(f"达到失败提醒间隔 ({self.params['fail_interval']}次)，但邮件功能未启用，跳过发送提醒")

            # 检查是否继续
            if not self.running:
                break

            # 等待下次查询
            self.log_signal.emit(f"{self.params['interval']}秒后再次尝试...")
            self.status_signal.emit(f"等待中... ({self.params['interval']}秒)")

            # 分多次等待，以便可以中断
            for i in range(self.params["interval"]):
                if not self.running:
                    break
                if i % 30 == 0:  # 每30秒记录一次等待状态
                    remaining = self.params["interval"] - i
                    self.status_signal.emit(f"等待中... ({remaining}秒)")
                time.sleep(1)

        if self.running:
            if self.attempt_count >= self.params["max_attempts"] and self.params["max_attempts"] > 0:
                self.log_signal.emit("\n已达到最大尝试次数，停止查询")
                self.status_signal.emit("已达最大尝试次数")
            else:
                self.log_signal.emit("\n查询已停止")
                self.status_signal.emit("已停止")

    def stop(self):
        """停止线程"""
        self.running = False
        self.log_signal.emit("正在停止查询...")

    def query_score(self):
        """查询软考成绩"""
        # 准备请求头
        headers = {
            "Accept": "*/*",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Cookie": self.params["cookie"],
            "Origin": "https://bm.ruankao.org.cn",
            "Referer": "https://bm.ruankao.org.cn/index.php/query/score",
            "Priority": "u=1, i",
            "Sec-Ch-Ua": '"Google Chrome";v="137", "Chromium";v="137", "Not/A)Brand";v="24"',
            "Sec-Ch-Ua-Mobile": "?0",
            "Sec-Ch-Ua-Platform": '"Windows"',
            "Sec-Fetch-Dest": "empty",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Site": "same-origin",
            "User-Agent": self.params["user_agent"],
            "X-Requested-With": "XMLHttpRequest"
        }

        # 准备表单数据（URL编码考试阶段）
        data = {
            "stage": self.params["stage"],
            "jym": ""  # 空验证码
        }

        try:
            self.log_signal.emit(f"请求URL: {self.params['url']}")
            self.log_signal.emit(f"请求数据: stage={data['stage']}, jym={data['jym']}")

            response = requests.post(
                self.params["url"],
                headers=headers,
                data=data,
                timeout=15
            )

            self.log_signal.emit(f"响应状态码: {response.status_code}")

            if response.status_code == 200:
                result = response.json()
                self.log_signal.emit(f"响应内容: {result}")

                # 检查是否包含成绩数据
                if result.get("msg") == "ok" and result.get("data"):
                    return True, result["data"]
                return False, f"成绩未公布: {result.get('msg', '未知状态')}"

            return False, f"HTTP错误: {response.status_code}"

        except Exception as e:
            return False, f"请求异常: {str(e)}"

    def send_email(self, subject, content):
        """发送邮件"""
        try:
            # 构建邮件内容
            full_content = f"""
            <h2>软考成绩查询通知</h2>
            <p>考试阶段：{self.params['stage']}</p>
            <p>查询时间：{time.strftime('%Y-%m-%d %H:%M:%S')}</p>
            <hr>
            {content}
            <p style="color:gray">此邮件由自动查询脚本发送，请勿回复</p>
            """

            msg = MIMEText(full_content, "html", "utf-8")
            msg["Subject"] = Header(subject, "utf-8")
            msg["From"] = self.params["sender_email"]
            msg["To"] = self.params["receiver_email"]

            with smtplib.SMTP_SSL(
                    self.params["smtp_server"],
                    self.params["smtp_port"]
            ) as server:
                server.login(self.params["sender_email"], self.params["sender_pwd"])
                server.sendmail(
                    self.params["sender_email"],
                    [self.params["receiver_email"]],
                    msg.as_string()
                )

            self.log_signal.emit(f"邮件发送成功: {subject}")
            return True
        except Exception as e:
            self.log_signal.emit(f"邮件发送失败: {str(e)}")
            return False


# 运行应用
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # 设置应用图标（可选）
    try:
        app.setWindowIcon(QIcon("智能体-1种尺寸.ico"))
    except:
        pass

    window = ScoreCheckerApp()
    window.show()
    sys.exit(app.exec_())
