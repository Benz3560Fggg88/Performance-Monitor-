import sys
import psutil
import time
import threading
import csv
import os
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QFileDialog, QHBoxLayout, QDoubleSpinBox, QCheckBox,
    QTableWidget, QTableWidgetItem, QSplitter, QHeaderView
)
from PyQt5.QtCore import Qt, pyqtSignal, QObject
from openpyxl import Workbook
from matplotlib.backends.backend_qt5agg import (
    FigureCanvasQTAgg as FigureCanvas,
    NavigationToolbar2QT as NavigationToolbar
)
from matplotlib.figure import Figure

# Worker class for emitting signals to the main thread
class Worker(QObject):
    update_ui = pyqtSignal(list, str)
    finish_monitoring_signal = pyqtSignal(str)

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
        self.ax_ram.set_xlabel("Time (H:MM:SS)")
        self.ax_ram.grid(True)
        self.figure.tight_layout(rect=[0, 0.03, 1, 0.95])

    def plot(self, timestamps, cpu_vals, ram_vals, is_real_time=True):
        self.ax_cpu.clear()
        self.ax_ram.clear()

        # Handle plotting for long-running processes in real-time mode
        if is_real_time and len(timestamps) > 1000:
            timestamps = timestamps[-1000:]
            cpu_vals = cpu_vals[-1000:]
            ram_vals = ram_vals[-1000:]

        self.ax_cpu.plot(timestamps, cpu_vals, '-', label='CPU (%)', color='tab:blue')
        self.ax_cpu.set_ylabel("CPU Usage (%)")
        self.ax_cpu.grid(True)
        self.ax_ram.plot(timestamps, ram_vals, '-', label='RAM (MB)', color='tab:orange')
        self.ax_ram.set_ylabel("RAM Usage (MB)")
        self.ax_ram.set_xlabel("Time (H:MM:SS)")
        self.ax_ram.grid(True)

        self.figure.suptitle("CPU and RAM Usage Over Time")
        self.figure.tight_layout(rect=[0, 0.03, 1, 0.95])
        self.canvas.draw()

    def reset_graph(self):
        self.ax_cpu.clear()
        self.ax_ram.clear()
        self.ax_cpu.grid(True)
        self.ax_ram.grid(True)
        self.figure.suptitle("CPU and RAM Usage Over Time")
        self.ax_cpu.set_ylabel("CPU Usage (%)")
        self.ax_ram.set_ylabel("RAM Usage (MB)")
        self.ax_ram.set_xlabel("Time (H:MM:SS)")
        self.canvas.draw()

class MonitorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("CPU/RAM Monitor by psutil")
        self.resize(1100, 700)

        temp_dir = "C:\\temp"
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
        self.auto_save_path = None
        self.worker = Worker()
        self.worker.update_ui.connect(self.update_ui)
        self.worker.finish_monitoring_signal.connect(self.finish_monitoring)

        self.table = QTableWidget(0, 4)
        self.table.setWordWrap(False)
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

        # --- New UI Components ---
        self.sampling_spinbox = QDoubleSpinBox()
        self.sampling_spinbox.setRange(0.1, 10.0)
        self.sampling_spinbox.setValue(1.0)
        self.sampling_spinbox.setSingleStep(0.1)

        self.auto_start_checkbox = QCheckBox("Start Detection Automatically")
        self.enable_plot_checkbox = QCheckBox("Enable Plotting")
        self.enable_plot_checkbox.setChecked(True)
        self.enable_plot_checkbox.stateChanged.connect(self.toggle_plot_options)

        self.plot_mode_checkbox = QCheckBox("Plot Graph After Training Ends")
        self.plot_mode_checkbox.setChecked(False)

        self.buffer_mode_checkbox = QCheckBox("Mode (tick=real-time, untick=buffered)")
        self.buffer_mode_checkbox.setChecked(False)

        self.btn_reset = QPushButton("Reset Table")
        self.btn_export_excel = QPushButton("Export to Excel")
        self.btn_export_csv = QPushButton("Export to CSV")
        self.btn_save_graph = QPushButton("Save Graph")
        self.btn_exit = QPushButton("Exit")
        
        self.btn_select_autosave = QPushButton("Select Auto-Save File")
        self.auto_save_file_label = QLabel("No file selected")
        self.btn_select_autosave.clicked.connect(self.select_autosave_file)

        self.btn_reset.clicked.connect(self.reset_table)
        self.btn_export_excel.clicked.connect(self.export_excel)
        self.btn_export_csv.clicked.connect(self.export_csv)
        self.btn_save_graph.clicked.connect(self.save_graph)
        self.btn_exit.clicked.connect(self.close)

        self.graph = PlotCanvas(self)
        self.setup_ui()
        self.toggle_plot_options()
        threading.Thread(target=self.monitor_loop, daemon=True).start()

    def format_duration(self, seconds):
        try:
            s_int = int(seconds)
            milliseconds = int((seconds - s_int) * 1000)
            hours, remainder = divmod(s_int, 3600)
            minutes, secs = divmod(remainder, 60)
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
        control_layout.addWidget(self.btn_select_autosave)
        control_layout.addWidget(self.auto_save_file_label)
        control_layout.addWidget(self.btn_exit)

        checkbox_layout = QHBoxLayout()
        checkbox_layout.addWidget(self.auto_start_checkbox)
        checkbox_layout.addWidget(self.enable_plot_checkbox)
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

    def toggle_plot_options(self):
        is_enabled = self.enable_plot_checkbox.isChecked()
        self.plot_mode_checkbox.setVisible(is_enabled)
        # self.buffer_mode_checkbox.setVisible(is_enabled) # ไม่ซ่อนปุ่ม buffer mode
        
        # Keep the graph widget visible, but reset its content when plotting is disabled
        if not is_enabled:
            self.graph.reset_graph()

    def select_autosave_file(self):
        if self.monitoring:
            self.status_label.setText("Cannot change auto-save file while monitoring.")
            return

        path, _ = QFileDialog.getSaveFileName(self, "Select Auto-Save File", "", "Excel Files (*.xlsx);;CSV Files (*.csv)")
        if path:
            self.auto_save_path = path
            self.auto_save_file_label.setText(os.path.basename(path))
            self.status_label.setText(f"Auto-save file selected: {os.path.basename(path)}")
        else:
            self.auto_save_path = None
            self.auto_save_file_label.setText("No file selected")
            self.status_label.setText("Auto-save file selection cancelled.")

    def reset_table(self):
        self.data.clear()
        self.buffered_data.clear()
        self.table.setRowCount(0)
        self.graph.reset_graph()
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
            self.worker.finish_monitoring_signal.emit("Process terminated.")
            return None, None
        except Exception as e:
            print(f"Error getting process resource: {e}")
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

        if self.enable_plot_checkbox.isChecked():
            is_real_time_mode = self.buffer_mode_checkbox.isChecked()
            timestamps = [d[0] for d in self.data]
            cpu_vals = [d[1] for d in self.data]
            ram_vals = [d[2] for d in self.data]
            self.graph.plot(timestamps, cpu_vals, ram_vals, is_real_time_mode)
            
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
                        self.worker.finish_monitoring_signal.emit("Process not found.")
                        proc_obj = None
                else:
                    time.sleep(0.5)
                    continue

            if not psutil.pid_exists(self.training_pid):
                self.worker.finish_monitoring_signal.emit("Process terminated.")
                proc_obj = None
                continue

            start_of_loop = time.time()
            cpu, ram = self.get_training_process_resource(proc_obj)

            if cpu is not None:
                elapsed_time = time.time() - self.training_start_time
                self.buffered_data.append((elapsed_time, cpu, ram, self.training_source))

                # --- Auto-Save Logic ---
                if elapsed_time >= 3600.0 and len(self.buffered_data) > 0:
                    self.auto_save_data()

                # --- UI Update Logic ---
                is_real_time_mode = self.buffer_mode_checkbox.isChecked()
                if is_real_time_mode:
                    self.worker.update_ui.emit(self.buffered_data, "flush")
                else:
                    elapsed = time.time() - self.training_start_time
                    if not self.initial_buffer_flushed and elapsed >= 10:
                        self.worker.update_ui.emit(self.buffered_data, "flush")
                        self.last_update_time = time.time()
                        self.initial_buffer_flushed = True
                    elif self.initial_buffer_flushed:
                        self.update_interval = self.get_dynamic_update_interval(elapsed)
                        if time.time() - self.last_update_time >= self.update_interval:
                            self.worker.update_ui.emit(self.buffered_data, "flush")
                            self.last_update_time = time.time()

            time_spent = time.time() - start_of_loop
            sleep_time = max(0, self.sampling_rate - time_spent)
            time.sleep(sleep_time)

    def auto_save_data(self):
        try:
            path = self.auto_save_path
            if not path:
                downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                path = os.path.join(downloads_path, f"Data_{timestamp}.xlsx")
            
            if path.lower().endswith('.xlsx'):
                wb = Workbook()
                ws = wb.active
                ws.append(["Time (H:MM:SS.ms)", "CPU (%)", "RAM (MB)", "Source"])
                for row_data in self.data + self.buffered_data:
                    formatted_row = [self.format_duration(row_data[0])] + list(row_data[1:])
                    ws.append(formatted_row)
                ws.append([])
                ws.append(["", "", "", f"Command/Source: {self.training_source}"])
                wb.save(path)
            elif path.lower().endswith('.csv'):
                with open(path, mode='w', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerow(["Time (H:MM:SS.ms)", "CPU (%)", "RAM (MB)", "Source"])
                    for row_data in self.data + self.buffered_data:
                        formatted_row = [self.format_duration(row_data[0])] + list(row_data[1:])
                        writer.writerow(formatted_row)
                    writer.writerow([])
                    writer.writerow(["", "", "", f"Command/Source: {self.training_source}"])

            self.status_label.setText(f"Auto-saved data to {os.path.basename(path)} and reset.")
            self.reset_data_after_save()

        except Exception as e:
            self.status_label.setText(f"Error during auto-save: {e}")

    def reset_data_after_save(self):
        self.data.clear()
        self.buffered_data.clear()
        self.table.setRowCount(0)
        self.training_start_time = time.time()
        self.last_update_time = self.training_start_time

    def update_ui(self, new_data, action):
        if action == "flush":
            self.flush_buffer_to_table_and_graph()
        
    def finish_monitoring(self, message):
        self.monitoring = False
        self.sampling_spinbox.setEnabled(True)
        self.btn_select_autosave.setEnabled(True)
        self.status_label.setText(f"Status: {message}. Showing final result...")
        self.source_label.setText(f"Finished monitoring: {self.training_source}")
        
        # Flush any remaining data to show the final result
        self.flush_buffer_to_table_and_graph()
        
        if self.enable_plot_checkbox.isChecked() and self.plot_mode_checkbox.isChecked():
            timestamps = [d[0] for d in self.data]
            cpu_vals = [d[1] for d in self.data]
            ram_vals = [d[2] for d in self.data]
            self.graph.plot(timestamps, cpu_vals, ram_vals, is_real_time=False)
        
    def start_monitoring(self):
        self.sampling_rate = self.sampling_spinbox.value()
        self.monitoring = True
        self.reset_table()
        self.training_start_time = time.time()
        self.last_update_time = self.training_start_time
        self.initial_buffer_flushed = False
        self.idle_start_time = None
        
        # Disable controls during monitoring
        self.sampling_spinbox.setEnabled(False)
        self.btn_select_autosave.setEnabled(False)
        
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
                writer.writerow(["Time (H:MM:SS.ms)", "CPU (%)", "RAM (MB)", "Source"])
                for row in self.data:
                    formatted_row = [self.format_duration(row[0])] + list(row[1:])
                    writer.writerow(formatted_row)
                writer.writerow([])
                writer.writerow(["", "", "", f"Command/Source: {self.training_source}"])
            self.status_label.setText(f"Status: CSV saved to {path}")

    def save_graph(self):
        if not self.enable_plot_checkbox.isChecked():
            self.status_label.setText("Status: Graph plotting is disabled.")
            return

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