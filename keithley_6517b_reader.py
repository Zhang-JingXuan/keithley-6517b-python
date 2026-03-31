import sys
import os
import time
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout,
                             QLabel, QPushButton, QComboBox, QCheckBox, QWidget, QGridLayout)
from PyQt5.QtCore import QTimer, QThread, pyqtSignal, Qt
from serial.tools import list_ports
import serial
import matplotlib
matplotlib.use('Qt5Agg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure
import csv
from datetime import datetime

# 设置matplotlib使用英文字体
plt.rcParams['font.sans-serif'] = ['Arial']
plt.rcParams['axes.unicode_minus'] = False

# ================== 配置数据保存根目录 ==================
BASE_DATA_DIR = r"D:/experiment_data"  # 你可以修改为其他路径，例如 "./data" 或 "C:/MyData"

# ================== 串口通信线程（查询模式） ==================
class SerialThread(QThread):
    data_received = pyqtSignal(str)

    def __init__(self, serial_port, parent=None):
        super().__init__(parent)
        self.serial_port = serial_port
        self.running = True
        self.interval = 0.1  # 查询间隔（秒）

    def run(self):
        print("Serial thread started (query mode)")
        # 初始化仪器：设置为电阻测量，输出纯数值
        try:
            self.serial_port.write(b':FUNC "RES"\r')
            time.sleep(0.1)
            self.serial_port.write(b':FORM:ELEM READ\r')
            time.sleep(0.1)
        except Exception as e:
            print(f"Failed to initialize instrument: {e}")

        while self.running and self.serial_port.isOpen():
            try:
                # 发送查询命令
                self.serial_port.write(b':FETCh?\r')
                # 读取一行响应
                data = self.serial_port.readline().decode().strip()
                if data:
                    self.data_received.emit(data)
                # 等待下一个查询周期
                time.sleep(self.interval)
            except Exception as e:
                print(f"Serial read error: {e}")
                break

    def stop(self):
        self.running = False
        self.wait()

# ================== 主窗口 ==================
class SerialPlotter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Real-time Resistance Plotter")
        self.resize(900, 600)

        # 数据存储（保留所有数据，无点数限制）
        self.timestamps = []
        self.data_values = []
        self.start_time = None

        # 串口相关
        self.serial_port = None
        self.serial_thread = None
        self.serial_open = False
        self.current_csv_filename = None

        # 坐标类型标志
        self.log_y = False

        # 创建界面
        self.setup_ui()

        # 定时器：更新绘图
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_plot)
        self.timer.start(100)

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 控制区布局
        control_layout = QGridLayout()

        # 串口选择
        control_layout.addWidget(QLabel("Port:"), 0, 0)
        self.port_combo = QComboBox()
        self.refresh_ports()
        control_layout.addWidget(self.port_combo, 0, 1)

        # 刷新按钮
        self.refresh_btn = QPushButton("Refresh")
        self.refresh_btn.clicked.connect(self.refresh_ports)
        control_layout.addWidget(self.refresh_btn, 0, 2)

        # 波特率
        control_layout.addWidget(QLabel("Baud rate:"), 1, 0)
        self.baud_combo = QComboBox()
        self.baud_combo.addItems(["9600", "19200", "38400", "115200"])
        control_layout.addWidget(self.baud_combo, 1, 1)

        # 数据位
        control_layout.addWidget(QLabel("Data bits:"), 2, 0)
        self.data_bits_combo = QComboBox()
        self.data_bits_combo.addItems(["8", "7"])
        control_layout.addWidget(self.data_bits_combo, 2, 1)

        # 停止位
        control_layout.addWidget(QLabel("Stop bits:"), 3, 0)
        self.stop_bits_combo = QComboBox()
        self.stop_bits_combo.addItems(["1", "1.5", "2"])
        control_layout.addWidget(self.stop_bits_combo, 3, 1)

        # 对数坐标复选框
        self.log_checkbox = QCheckBox("Log Y-axis")
        self.log_checkbox.stateChanged.connect(self.on_log_scale_changed)
        control_layout.addWidget(self.log_checkbox, 4, 0)

        # 开始/停止按钮
        self.start_btn = QPushButton("Start")
        self.start_btn.clicked.connect(self.toggle_serial)
        control_layout.addWidget(self.start_btn, 4, 1, 1, 2)

        main_layout.addLayout(control_layout)

        # 绘图区
        self.figure = Figure(figsize=(8, 4), dpi=100)
        self.canvas = FigureCanvas(self.figure)
        self.ax = self.figure.add_subplot(111)
        self.ax.set_xlabel("Time (s)")
        self.ax.set_ylabel("Resistance (Ω)")
        self.ax.grid(True)

        # 设置Y轴为科学计数法
        self.ax.ticklabel_format(axis='y', style='sci', scilimits=(0,0), useMathText=True)
        self.ax.yaxis.get_offset_text().set_fontsize(10)

        self.line, = self.ax.plot([], [], 'b-', linewidth=1)

        # 添加工具栏（支持缩放、平移、保存等）
        self.toolbar = NavigationToolbar(self.canvas, self)
        main_layout.addWidget(self.toolbar)
        main_layout.addWidget(self.canvas)

    def on_log_scale_changed(self, state):
        self.log_y = (state == Qt.Checked)
        self.update_plot()

    def refresh_ports(self):
        self.port_combo.clear()
        ports = [p.device for p in list_ports.comports()]
        self.port_combo.addItems(ports)

    def toggle_serial(self):
        if self.serial_open:
            self.stop_serial()
        else:
            self.start_serial()

    def start_serial(self):
        port_name = self.port_combo.currentText()
        if not port_name:
            print("Please select a port.")
            return

        baud_rate = int(self.baud_combo.currentText())
        data_bits = serial.EIGHTBITS if self.data_bits_combo.currentText() == "8" else serial.SEVENBITS
        stop_bits_text = self.stop_bits_combo.currentText()
        if stop_bits_text == "1":
            stop_bits = serial.STOPBITS_ONE
        elif stop_bits_text == "1.5":
            stop_bits = serial.STOPBITS_ONE_POINT_FIVE
        else:
            stop_bits = serial.STOPBITS_TWO

        try:
            self.serial_port = serial.Serial(
                port=port_name,
                baudrate=baud_rate,
                bytesize=data_bits,
                stopbits=stop_bits,
                timeout=0.5
            )
            # 清空之前的数据（开始新实验）
            self.timestamps.clear()
            self.data_values.clear()
            self.start_time = datetime.now()

            # 生成本次实验的 CSV 文件名
            now = datetime.now()
            date_str = now.strftime("%Y-%m-%d")
            time_str = now.strftime("%Y%m%d_%H%M%S")
            folder_path = os.path.join(BASE_DATA_DIR, date_str)
            os.makedirs(folder_path, exist_ok=True)
            self.current_csv_filename = os.path.join(folder_path, f"{time_str}.csv")

            # 初始化 CSV 文件，写入表头
            try:
                with open(self.current_csv_filename, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(['Time (s)', 'Resistance (Ω)', 'Raw data'])
                print(f"Data will be saved to: {self.current_csv_filename}")
            except Exception as e:
                print(f"Failed to create CSV file: {e}")
                self.current_csv_filename = None

            # 启动线程
            self.serial_thread = SerialThread(self.serial_port)
            self.serial_thread.data_received.connect(self.handle_data_received)
            self.serial_thread.start()

            self.serial_open = True
            self.start_btn.setText("Stop")
            print(f"Port {port_name} opened, thread started.")
        except Exception as e:
            print(f"Failed to open port: {e}")

    def stop_serial(self):
        if self.serial_thread is not None:
            self.serial_thread.stop()
            self.serial_thread = None
        if self.serial_port is not None and self.serial_port.isOpen():
            self.serial_port.close()
        self.serial_open = False
        self.start_btn.setText("Start")
        self.current_csv_filename = None
        print("Serial port closed.")

    def handle_data_received(self, raw_line):
        """处理接收到的每一行数据"""
        print(f"Raw: {raw_line}")  # 调试输出，可注释掉
        try:
            # 直接转换为浮点数（因为已设置为纯数值输出）
            value = float(raw_line)
            elapsed = (datetime.now() - self.start_time).total_seconds()
            # 存储所有数据
            self.timestamps.append(elapsed)
            self.data_values.append(value)

            # 保存到CSV
            self.save_to_csv(elapsed, value, raw_line)

        except Exception as e:
            print(f"Parse error: {e}")

    def save_to_csv(self, elapsed, value, raw_line):
        """保存数据到当前实验的CSV文件"""
        if self.current_csv_filename is None:
            return
        try:
            with open(self.current_csv_filename, 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow([f"{elapsed:.3f}", f"{value:.6e}", raw_line])
        except Exception as e:
            print(f"CSV save error: {e}")

    def update_plot(self):
        """更新绘图（绘制所有数据点）"""
        if self.timestamps and self.data_values:
            self.line.set_data(self.timestamps, self.data_values)
            # 根据对数复选框设置Y轴比例
            if self.log_y:
                if all(v > 0 for v in self.data_values):
                    self.ax.set_yscale('log')
                else:
                    print("Warning: Data contains non-positive values, cannot use log scale. Switching to linear.")
                    self.log_checkbox.setChecked(False)
                    self.log_y = False
                    self.ax.set_yscale('linear')
            else:
                self.ax.set_yscale('linear')
            # 自动调整坐标轴范围
            self.ax.relim()
            self.ax.autoscale_view()
            self.canvas.draw_idle()
        else:
            if self.log_y:
                self.ax.set_yscale('linear')
            self.canvas.draw_idle()

    def closeEvent(self, event):
        self.stop_serial()
        event.accept()

# ================== 程序入口 ==================
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = SerialPlotter()
    window.show()
    sys.exit(app.exec_())