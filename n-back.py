import sys
import os
import time
import random
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QComboBox, QPushButton, QWidget, QVBoxLayout, QHBoxLayout, QMessageBox, QTableWidget, QTableWidgetItem
from PyQt5.QtCore import Qt, QTimer
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import xlwings as xw
# 定义颜色常量
BG_COLOR_DEFAULT = "#2c3e50"
BG_COLOR_ACTIVE = "#3498db"
BG_COLOR_CORRECT = "#2ecc71"
BG_COLOR_WRONG = "#e74c3c"
BG_COLOR_WIN = "#ffffff"
# 设置matplotlib支持中文显示
plt.rcParams['font.family'] = 'SimHei'
plt.rcParams['axes.unicode_minus'] = False


class NBackTestApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.initialize_variables()

    def initUI(self):
        self.setWindowTitle("N-Back认知测试")
        self.setGeometry(600, 200, 800, 750)

        # 中心部件和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 顶部控制面板
        control_layout = QHBoxLayout()
        self.n_combo = QComboBox()
        self.n_combo.addItems(["1", "2", "3", "4", "5", "6"])
        self.length_combo = QComboBox()
        self.length_combo.addItems([str(i) for i in range(5, 30)])
        self.start_button = QPushButton("开始测试")
        self.start_button.clicked.connect(self.start_test)
        self.instruction_label = QLabel("说明: 当屏幕上的数字与N次之前出现的数字相同时，按 Z 键或点击“相同”按钮；不同时，按 M 键或点击“不同”按钮。N值可在下拉菜单中选择。测试开始的前 N 个数字无需操作。")
        self.instruction_label.setWordWrap(True)
        control_layout.addWidget(QLabel("N值:"))
        control_layout.addWidget(self.n_combo)
        control_layout.addWidget(QLabel("测试轮次:"))
        control_layout.addWidget(self.length_combo)
        control_layout.addWidget(self.start_button)
        main_layout.addWidget(self.instruction_label)
        main_layout.addLayout(control_layout)

        # 数据显示区域
        self.number_label = QLabel("", alignment=Qt.AlignCenter)
        self.number_label.setStyleSheet(f"font-size: 72px; background-color: {BG_COLOR_DEFAULT}; color: white; border-radius: 10px; padding: 20px;")
        main_layout.addWidget(self.number_label)

        # 倒计时提示
        self.countdown_label = QLabel("", alignment=Qt.AlignCenter)
        self.countdown_label.setStyleSheet("font-size: 24px; color: black;")
        main_layout.addWidget(self.countdown_label)

        # 统计面板
        stats_layout = QHBoxLayout()
        self.acc_label = QLabel("当前正确率: 0%")
        self.total_tests_label = QLabel("总测试次数: 0")
        stats_layout.addWidget(self.acc_label)
        stats_layout.addWidget(self.total_tests_label)
        main_layout.addLayout(stats_layout)

        # 新增按钮布局
        button_group_layout = QHBoxLayout()
        self.same_button = QPushButton("相同")
        self.same_button.clicked.connect(lambda: self.check_answer_by_button(True))
        self.different_button = QPushButton("不同")
        self.different_button.clicked.connect(lambda: self.check_answer_by_button(False))
        button_group_layout.addWidget(self.same_button)
        button_group_layout.addWidget(self.different_button)
        main_layout.addLayout(button_group_layout)

        # 底部按钮
        button_layout = QHBoxLayout()
        self.export_button = QPushButton("导出数据")
        self.export_button.clicked.connect(self.export_data)
        self.exit_button = QPushButton("退出")
        self.exit_button.clicked.connect(self.close)
        button_layout.addWidget(self.export_button)
        button_layout.addWidget(self.exit_button)
        main_layout.addLayout(button_layout)

        # 图表区域
        self.figure = Figure(figsize=(6, 4))
        self.canvas = FigureCanvas(self.figure)
        main_layout.addWidget(self.canvas)

        # 初始化图表
        self.ax = self.figure.add_subplot(111)
        self.ax.set_ylabel('正确率 (%)')
        self.ax.set_title('历史正确率趋势')
        self.ax.set_xticks([])

        # 整体窗口样式
        # self.setStyleSheet(f"background-color: {BG_COLOR_WIN}; color: black;")

    def initialize_variables(self):
        self.test_active = False
        self.sequence = []
        self.current_index = 0
        self.correct = 0
        self.total = 0
        self.history = []
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_display)
        self.response_received = False
        self.countdown = 3
        self.countdown_timer = QTimer()
        self.countdown_timer.timeout.connect(self.update_countdown)
        self.key_pressed = False

    def generate_sequence(self):
        self.length = int(self.length_combo.currentText()) + int(self.n_combo.currentText())
        self.sequence = [random.randint(1, 9) for _ in range(self.length)]
        n = int(self.n_combo.currentText())
        for i in range(n, len(self.sequence)):
            if random.random() < 0.3:
                self.sequence[i] = self.sequence[i - n]

    def start_test(self):
        if self.test_active:
            return
        self.test_active = True
        self.generate_sequence()
        self.current_index = 0
        self.correct = 0
        self.total = 0
        self.history = []
        self.update_display()
        self.setFocus()
        self.keyPressEvent = self.check_answer

    def update_countdown(self):
        self.countdown -= 1
        if self.countdown >= 0:
            self.countdown_label.setText(f"倒计时: {self.countdown}")
        else:
            self.countdown_timer.stop()
            n = int(self.n_combo.currentText())
            if self.current_index < n:
                self.total = 0
                self.correct = 0
                self.update_stats()
                self.current_index += 1
            else:
                if not self.response_received:
                    self.total += 1
                    self.process_no_response()
            self.number_label.setStyleSheet(f"font-size: 72px; background-color: {BG_COLOR_DEFAULT}; color: white; border-radius: 10px; padding: 20px;")
            self.update_display()

    def check_answer(self, event):
        n = int(self.n_combo.currentText())
        if not self.test_active or self.current_index < n or self.key_pressed:
            return
        correct_answer = (self.sequence[self.current_index] == self.sequence[self.current_index - n])
        if (event.key() == Qt.Key_Z and correct_answer) or (event.key() == Qt.Key_M and not correct_answer):
            self.handle_correct_answer()
        else:
            self.handle_wrong_answer()
        self.response_received = True
        self.total += 1
        self.update_stats()
        if self.current_index < len(self.sequence):
            self.number_label.setText(str(self.sequence[self.current_index]))
        self.timer.start(500)
        self.countdown_timer.stop()
        self.key_pressed = True

    def check_answer_by_button(self, is_same):
        n = int(self.n_combo.currentText())
        if not self.test_active or self.current_index < n or self.key_pressed:
            return
        correct_answer = (self.sequence[self.current_index] == self.sequence[self.current_index - n])
        if (is_same and correct_answer) or (not is_same and not correct_answer):
            self.handle_correct_answer()
        else:
            self.handle_wrong_answer()
        self.response_received = True
        self.total += 1
        self.update_stats()
        if self.current_index < len(self.sequence):
            self.number_label.setText(str(self.sequence[self.current_index]))
        self.timer.start(500)
        self.countdown_timer.stop()
        self.key_pressed = True

    def update_display(self):
        self.timer.stop()
        if self.current_index >= len(self.sequence):
            self.end_test()
            return
        self.number_label.setText(str(self.sequence[self.current_index]))
        self.number_label.setStyleSheet(f"font-size: 72px; background-color: {BG_COLOR_ACTIVE}; color: white; border-radius: 10px; padding: 20px;")
        self.response_received = False
        self.countdown = 2
        self.countdown_label.setText(f"倒计时: {self.countdown}")
        self.countdown_timer.start(1000)
        self.key_pressed = False

    def handle_correct_answer(self):
        self.correct += 1
        self.is_correct = 1
        self.current_index += 1
        self.number_label.setStyleSheet(f"font-size: 72px; background-color: {BG_COLOR_CORRECT}; color: white; border-radius: 10px; padding: 20px;")

    def handle_wrong_answer(self):
        self.current_index += 1
        self.is_correct = 0
        self.number_label.setStyleSheet(f"font-size: 72px; background-color: {BG_COLOR_WRONG}; color: white; border-radius: 10px; padding: 20px;")

    def process_no_response(self):
        n = int(self.n_combo.currentText())
        if self.current_index >= n and not self.response_received:
            self.update_stats()
        self.current_index += 1
        self.number_label.setStyleSheet(f"font-size: 72px; background-color: {BG_COLOR_DEFAULT}; color: white; border-radius: 10px; padding: 20px;")
        self.update_display()

    def update_stats(self):
        if self.current_index > int(self.n_combo.currentText()):
            acc = round((self.correct / self.total * 100), 2)
            self.acc_label.setText(f"正确率: {acc:.2f}%")
            self.history.append({'n_value': int(self.n_combo.currentText()), 'is_correct': self.is_correct, 'Test_rounds': self.total,
                                 'accuracy': acc, 'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
        total_tests = self.total
        self.total_tests_label.setText(f"总测试次数: {total_tests}")

    def end_test(self):
        self.test_active = False
        self.timer.stop()
        self.keyPressEvent = self.default_key_event
        if self.total > 0:
            acc = self.correct / self.total * 100
        else:
            acc = 0
        self.plot_history()
        self.show_custom_message_box("测试完成", f"测试结束！正确率: {acc:.2f}%")

    def plot_history(self):
        self.ax.clear()
        df = pd.DataFrame(self.history)
        if not df.empty:
            print(df)
            df['timestamp'] = pd.to_datetime(df['timestamp'])
            df = df.sort_values('timestamp')
            marker = ['.', 'o', 'v', '^', '<', '>', '1', '2', '3', '4', '8', 's', 'p', '*', 'H', '+', 'x', 'd']
            marker_size = '12'
            self.ax.plot(df['timestamp'], df['accuracy'], marker=marker[8], ms=marker_size)
            self.ax.set_ylabel('正确率 (%)')
            self.ax.set_title('历史正确率趋势')
            self.ax.tick_params(axis='x', rotation=5)
        self.canvas.draw()

    def export_data(self):
        if not self.history:
            self.show_custom_message_box("无数据", "没有可导出的测试数据")
            return
        df = pd.DataFrame(self.history)
        try:
            try:
                os.mkdir('./Excel')
            except Exception as e:
                pass
            file_name = "n-back" + time.strftime('%Y.%m.%d %H:%M:%S ',
                                                 time.localtime(time.time())).replace(":", "-") + ".xlsx"
            export_path = os.getcwd() + '\\' + 'Excel'
            df.to_excel(export_path + '\\' + file_name, index=False)
            file_path = os.path.join(export_path, file_name)
            msg_box = QMessageBox()
            msg_box.setWindowTitle("导出提示")
            msg_box.setText(f"导出成功!\n文件路径:{file_path}")
            msg_box.addButton("立刻打开", QMessageBox.YesRole)
            msg_box.addButton("关闭", QMessageBox.NoRole)
            msg_box.setStyleSheet("QLabel{min-width: 400px; font-size: 16px;} QPushButton{min-width: 100px; font-size: 16px;}")
            msg_box.exec_()
            try:
                if msg_box.clickedButton().text() == "立刻打开":
                    print("打开中")
                    app = xw.App(visible=True, add_book=False)
                    app.display_alerts = False
                    # 文件位置,打开excel文档
                    app.books.open(file_path)
            except Exception as e:
                self.show_custom_message_box("打开失败", f"打开文件失败: {str(e)}")
            else:
                pass
        except Exception as e:
            self.show_custom_message_box("导出失败", f"导出失败!{str(e)}")

    def show_custom_message_box(self, title, message):
        msg_box = QMessageBox()
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setStyleSheet("QLabel{min-width: 400px; font-size: 16px;} QPushButton{min-width: 100px; font-size: 16px;}")
        msg_box.exec_()

    def default_key_event(self, event):
        pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = NBackTestApp()
    ex.show()
    sys.exit(app.exec_())