import sys, psutil, time, threading, csv, os
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QFileDialog, QHBoxLayout, QDoubleSpinBox, QCheckBox,
    QTableWidget, QTableWidgetItem, QSplitter, QHeaderView
)
from PyQt5.QtCore import Qt
from openpyxl import Workbook
from matplotlib.backends.backend_qt5agg import (
    FigureCanvasQTAgg as FigureCanvas,
    NavigationToolbar2QT as NavigationToolbar
)
from matplotlib.figure import Figure

class PlotCanvas(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.figure = Figure()
        self.ax_cpu, self.ax_ram = self.figure.subplots(2, 1, sharex=True)

        self.canvas = FigureCanvas(self.figure)
        self.toolbar = NavigationToolbar(self.canvas, self)

        layout = QVBoxLayout()
        layout.addWidget(self.toolbar)
        layout.addWidget(self.canvas)
        self.setLayout(layout)

        self.figure.suptitle("CPU and RAM Usage Over Time")
        self.ax_cpu.set_ylabel("CPU Usage (%)")
        self.ax_cpu.grid(True)
        self.ax_ram.set_ylabel("RAM Usage (MB)")
        self.ax_ram.set_xlabel("Time")
        self.ax_ram.grid(True)

        self.figure.tight_layout(rect=[0, 0.03, 1, 0.95])

    def plot(self, timestamps, cpu_vals, ram_vals):
        self.ax_cpu.clear()
        self.ax_ram.clear()

        self.ax_cpu.plot(timestamps, cpu_vals, '-', label='CPU (%)', color='tab:blue')
        self.ax_cpu.set_ylabel("CPU Usage (%)")
        self.ax_cpu.grid(True)

        self.ax_ram.plot(timestamps, ram_vals, '-', label='RAM (MB)', color='tab:orange')
        self.ax_ram.set_ylabel("RAM Usage (MB)")
        self.ax_ram.set_xlabel("Time")
        self.ax_ram.grid(True)

        self.figure.suptitle("CPU and RAM Usage Over Time")

        self.figure.tight_layout(rect=[0, 0.03, 1, 0.95])
        self.canvas.draw()

class MonitorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CPU/RAM Monitor by psutil")
        self.resize(1100, 700)

        temp_dir = "C:\\temp"                                    # สร้างtemp กรณีไม่มี
        if not os.path.exists(temp_dir):                         
            os.makedirs(temp_dir)  

        self.monitoring = False
        self.training_source = "Manual"
        self.training_pid = None
        self.data = []
        self.buffered_data = []
        self.sampling_rate = 1.0
        self.training_start_time = None
        self.last_update_time = time.time()
        self.update_interval = 2
        self.initial_buffer_flushed = False
        self.idle_start_time = None
        self.IDLE_THRESHOLD_SECONDS = 30

        self.table = QTableWidget(0, 4)
        self.table.setWordWrap(False)
        # --- เปลี่ยนชื่อหัวตาราง ---
        self.table.setHorizontalHeaderLabels(["Time (H:MM:SS.ms)", "CPU (%)", "RAM (MB)", "Source"])
        self.table.verticalHeader().setVisible(True)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)

        self.status_label = QLabel("Status: Idle")
        self.source_label = QLabel("")
        self.sampling_spinbox = QDoubleSpinBox()
        self.sampling_spinbox.setRange(0.1, 10.0)
        self.sampling_spinbox.setValue(1.0)
        self.sampling_spinbox.setSingleStep(0.1)
        self.auto_start_checkbox = QCheckBox("Start Detection Automatically")
        self.plot_mode_checkbox = QCheckBox("Plot Graph After Training Ends")
        self.buffer_mode_checkbox = QCheckBox("Mode (tick=real-time, untick=buffered)")
        self.buffer_mode_checkbox.setChecked(False)
        self.btn_reset = QPushButton("Reset Table")
        self.btn_export_excel = QPushButton("Export to Excel")
        self.btn_export_csv = QPushButton("Export to CSV")
        self.btn_save_graph = QPushButton("Save Graph")
        self.btn_exit = QPushButton("Exit")
        self.btn_reset.clicked.connect(self.reset_table)
        self.btn_export_excel.clicked.connect(self.export_excel)
        self.btn_export_csv.clicked.connect(self.export_csv)
        self.btn_save_graph.clicked.connect(self.save_graph)
        self.btn_exit.clicked.connect(self.close)
        self.graph = PlotCanvas(self)
        self.setup_ui()
        threading.Thread(target=self.monitor_loop, daemon=True).start()

    # *** แก้ไขจุดที่ 1: เปลี่ยนฟังก์ชัน format_duration ทั้งหมด ***
    def format_duration(self, seconds):
        """Converts seconds into H:MM:SS.ms format."""
        try:
            # แยกส่วนจำนวนเต็ม (วินาที) และส่วนทศนิยม (สำหรับมิลลิวินาที)
            s_int = int(seconds)
            milliseconds = int((seconds - s_int) * 1000)

            # คำนวณชั่วโมงรวม นาที และวินาที จากส่วนจำนวนเต็ม
            hours, remainder = divmod(s_int, 3600)
            minutes, secs = divmod(remainder, 60)

            # จัดรูปแบบให้อยู่ใน H:MM:SS.ms
            return f"{hours}:{minutes:02d}:{secs:02d}.{milliseconds:03d}"
        except (ValueError, TypeError):
            return str(seconds)

    def setup_ui(self):
        layout = QVBoxLayout()
        control_layout = QHBoxLayout()
        control_layout.addWidget(QLabel("Sampling Rate (s):"))
        control_layout.addWidget(self.sampling_spinbox)
        control_layout.addStretch()
        control_layout.addWidget(self.btn_reset)
        control_layout.addWidget(self.btn_export_excel)
        control_layout.addWidget(self.btn_export_csv)
        control_layout.addWidget(self.btn_save_graph)
        control_layout.addWidget(self.btn_exit)
        checkbox_layout = QHBoxLayout()
        checkbox_layout.addWidget(self.auto_start_checkbox)
        checkbox_layout.addWidget(self.plot_mode_checkbox)
        checkbox_layout.addWidget(self.buffer_mode_checkbox)
        checkbox_layout.addStretch()
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self.table)
        splitter.addWidget(self.graph)
        splitter.setSizes([400, 700])
        layout.addWidget(self.status_label)
        layout.addWidget(self.source_label)
        layout.addLayout(checkbox_layout)
        layout.addWidget(splitter)
        layout.addLayout(control_layout)
        self.setLayout(layout)

    def reset_table(self):
        self.data.clear()
        self.buffered_data.clear()
        self.table.setRowCount(0)
        self.graph.ax_cpu.clear()
        self.graph.ax_ram.clear()
        self.graph.figure.suptitle("CPU and RAM Usage Over Time")
        self.graph.ax_cpu.set_ylabel("CPU Usage (%)")
        self.graph.ax_cpu.grid(True)
        self.graph.ax_ram.set_ylabel("RAM Usage (MB)")
        self.graph.ax_ram.set_xlabel("Time")
        self.graph.ax_ram.grid(True)
        self.graph.figure.tight_layout(rect=[0, 0.03, 1, 0.95])
        self.graph.canvas.draw()
        self.status_label.setText("Status: Table and graph reset.")
        self.source_label.setText("")

    def detect_training_process(self):
        try:
            with open("C:\\temp\\training_pid.txt", "r") as f:
                pid = int(f.read().strip())
                proc = psutil.Process(pid)
                if proc.is_running():
                    cmd = ' '.join(proc.cmdline())
                    self.training_source = f"MATLAB (PID: {pid}) CMD: {cmd}"
                    self.training_pid = pid
                    return True
        except (FileNotFoundError, ValueError, psutil.NoSuchProcess):
            pass
        my_pid = psutil.Process().pid
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                if proc.pid == my_pid: continue
                cmdline_args = proc.info.get('cmdline')
                if not cmdline_args:
                    continue
                name = proc.info['name'].lower()
                if "python" in name and any(arg.endswith(".py") for arg in cmdline_args):
                    self.training_source = f"Python: {' '.join(proc.info['cmdline'])}"
                    self.training_pid = proc.pid
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
        return False

    def get_training_process_resource(self, proc):
        try:
            cpu = proc.cpu_percent(interval=None) / psutil.cpu_count()
            ram = proc.memory_info().rss / (1024 * 1024)
            return cpu, ram
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            self.finish_monitoring()
        except Exception:
            pass
        return None, None

    def flush_buffer_to_table_and_graph(self):
        if not self.buffered_data:
            return
        for rowdata in self.buffered_data:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setVerticalHeaderItem(row, QTableWidgetItem(str(row + 1)))
            for i, val in enumerate(rowdata):
                if i == 0:
                    item_text = self.format_duration(val)
                else:
                    item_text = f"{val:.2f}" if isinstance(val, float) else str(val)
                self.table.setItem(row, i, QTableWidgetItem(item_text))

        self.data.extend(self.buffered_data)
        if not self.plot_mode_checkbox.isChecked():
            timestamps = [d[0] for d in self.data]
            cpu_vals = [d[1] for d in self.data]
            ram_vals = [d[2] for d in self.data]
            self.graph.plot(timestamps, cpu_vals, ram_vals)
        self.buffered_data.clear()
        self.table.scrollToBottom()

    def get_dynamic_update_interval(self, elapsed_seconds):
        if elapsed_seconds <= 10: return 10
        if elapsed_seconds <= 20: return 2
        if elapsed_seconds <= 60: return 5
        if elapsed_seconds <= 300: return 10
        if elapsed_seconds <= 900: return 20
        return 30

    def monitor_loop(self):
        proc_obj = None
        while True:
            if not self.monitoring:
                if self.auto_start_checkbox.isChecked() and self.detect_training_process():
                    self.start_monitoring()
                    try:
                        proc_obj = psutil.Process(self.training_pid)
                        proc_obj.cpu_percent(interval=None)
                        time.sleep(self.sampling_rate)
                    except psutil.NoSuchProcess:
                        self.finish_monitoring()
                        proc_obj = None
                else:
                    time.sleep(0.5)
                    continue
            if not psutil.pid_exists(self.training_pid):
                self.finish_monitoring()
                proc_obj = None
                continue
            start_of_loop = time.time()
            cpu, ram = self.get_training_process_resource(proc_obj)
            if 'matlab' in self.training_source.lower():
                if not os.path.exists("C:\\temp\\training_pid.txt"):
                    self.finish_monitoring()
                    continue
                if cpu is not None and cpu < 1.0:
                    if self.idle_start_time is None:
                        self.idle_start_time = time.time()
                    elif time.time() - self.idle_start_time > self.IDLE_THRESHOLD_SECONDS:
                        self.finish_monitoring()
                        continue
                else:
                    self.idle_start_time = None
            if cpu is not None:
                elapsed_time = time.time() - self.training_start_time
                self.buffered_data.append((elapsed_time, cpu, ram, self.training_source))
            is_sampling_mode = self.buffer_mode_checkbox.isChecked()
            if is_sampling_mode:
                self.flush_buffer_to_table_and_graph()
            else:
                elapsed = time.time() - self.training_start_time
                if not self.initial_buffer_flushed and elapsed >= 10:
                    self.flush_buffer_to_table_and_graph()
                    self.last_update_time = time.time()
                    self.initial_buffer_flushed = True
                elif self.initial_buffer_flushed:
                    self.update_interval = self.get_dynamic_update_interval(elapsed)
                    if time.time() - self.last_update_time >= self.update_interval:
                        self.flush_buffer_to_table_and_graph()
                        self.last_update_time = time.time()
            time_spent = time.time() - start_of_loop
            sleep_time = max(0, self.sampling_rate - time_spent)
            time.sleep(sleep_time)

    def finish_monitoring(self):
        self.monitoring = False
        self.status_label.setText("Training stopped. Showing final result...")
        self.flush_buffer_to_table_and_graph()
        if self.plot_mode_checkbox.isChecked() or not self.data:
            timestamps = [d[0] for d in self.data]
            cpu_vals = [d[1] for d in self.data]
            ram_vals = [d[2] for d in self.data]
            self.graph.plot(timestamps, cpu_vals, ram_vals)
        self.source_label.setText(f"Finished monitoring: {self.training_source}")

    def start_monitoring(self):
        self.sampling_rate = self.sampling_spinbox.value()
        self.monitoring = True
        self.reset_table()
        self.training_start_time = time.time()
        self.last_update_time = self.training_start_time
        self.initial_buffer_flushed = False
        self.idle_start_time = None
        self.status_label.setText("Monitoring...")
        self.source_label.setText(f"Monitoring process: {self.training_source}")

    def export_excel(self):
        if not self.data:
            self.status_label.setText("Status: No data to export")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx)")
        if path:
            wb = Workbook()
            ws = wb.active
            # --- เปลี่ยนชื่อหัวตาราง ---
            ws.append(["Time (H:MM:SS.ms)", "CPU (%)", "RAM (MB)", "Source"])
            for row in self.data:
                formatted_row = [self.format_duration(row[0])] + list(row[1:])
                ws.append(formatted_row)
            ws.append([])
            ws.append(["", "", "", f"Command/Source: {self.training_source}"])
            wb.save(path)
            self.status_label.setText(f"Status: Excel saved to {path}")

    def export_csv(self):
        if not self.data:
            self.status_label.setText("Status: No data to export")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save CSV File", "", "CSV Files (*.csv)")
        if path:
            with open(path, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                # --- เปลี่ยนชื่อหัวตาราง ---
                writer.writerow(["Time (H:MM:SS.ms)", "CPU (%)", "RAM (MB)", "Source"])
                for row in self.data:
                    formatted_row = [self.format_duration(row[0])] + list(row[1:])
                    writer.writerow(formatted_row)
                writer.writerow([])
                writer.writerow(["", "", "", f"Command/Source: {self.training_source}"])
            self.status_label.setText(f"Status: CSV saved to {path}")

    def save_graph(self):
        if not self.data:
            self.status_label.setText("Status: No data to save graph")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save Graph as Image", "", "PNG Files (*.png)")
        if path:
            self.graph.figure.savefig(path, dpi=300, bbox_inches='tight')
            self.status_label.setText(f"Status: Graph saved to {path}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MonitorApp()
    win.show()
    sys.exit(app.exec_())