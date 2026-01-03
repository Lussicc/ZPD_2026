
"""
Piezoelektriskā lietus sensora programmatūra
ZPD autors: Eduards Dāvis Ziemelis, 2026
Jelgavas Valsts ģimnāzija

Šis kods tika pilnībā ģenerēts ar mākslīgā intelekta rīku Claude Sonnet 4.5 
(Anthropic), pamatojoties uz darba autora detalizētiem tehniskajiem 
aprakstiem par aparatūras konfigurāciju, datu apstrādes prasībām un 
vēlamajām funkcijām.
"""

import sys
import serial
import serial.tools.list_ports
from collections import deque
import numpy as np
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QComboBox, QLabel, 
                             QGroupBox, QCheckBox, QDoubleSpinBox,
                             QGridLayout, QScrollArea, QSplitter, QDialog,
                             QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox)
from PyQt5.QtCore import QTimer, pyqtSignal, QThread, Qt
from PyQt5.QtGui import QFont
import pyqtgraph as pg
from scipy import signal as scipy_signal

class SpikeResultsDialog(QDialog):
    """Dialog to display spike detection results"""
    def __init__(self, spike_data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Spike Detection Results")
        self.setGeometry(200, 200, 500, 400)
        
        layout = QVBoxLayout()
        
        # Info label
        info_label = QLabel(f"Found {len(spike_data)} spikes. Select cells and Ctrl+C to copy to Excel.")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        # Create table
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(['Nr', 'Time (s)', 'Max', 'Min'])
        self.table.setRowCount(len(spike_data))
        
        # Populate table
        for i, spike in enumerate(spike_data):
            self.table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
            self.table.setItem(i, 1, QTableWidgetItem(f"{spike['time']:.4f}"))
            self.table.setItem(i, 2, QTableWidgetItem(f"{spike['max']:.2f}"))
            self.table.setItem(i, 3, QTableWidgetItem(f"{spike['min']:.2f}"))
        
        # Make table look nice
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionMode(QTableWidget.ContiguousSelection)
        
        layout.addWidget(self.table)
        
        # Buttons
        btn_layout = QHBoxLayout()
        
        copy_btn = QPushButton("Copy All to Clipboard")
        copy_btn.clicked.connect(self.copy_all_to_clipboard)
        btn_layout.addWidget(copy_btn)
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        btn_layout.addWidget(close_btn)
        
        layout.addLayout(btn_layout)
        self.setLayout(layout)
    
    def copy_all_to_clipboard(self):
        """Copy entire table to clipboard in Excel-friendly format"""
        text = "Nr\tTime (s)\tMax\tMin\n"
        for row in range(self.table.rowCount()):
            row_data = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                row_data.append(item.text() if item else "")
            text += "\t".join(row_data) + "\n"
        
        QApplication.clipboard().setText(text)
        QMessageBox.information(self, "Copied", "Table copied to clipboard! You can paste it in Excel now.")


class SerialReader(QThread):
    """Background thread for reading serial data"""
    data_received = pyqtSignal(np.ndarray, float)
    status_message = pyqtSignal(str)
    connection_changed = pyqtSignal(bool)
    
    def __init__(self):
        super().__init__()
        self.serial_port = None
        self.running = False
        self.port_name = None
        self.baud_rate = 921600
        
    def connect_to_port(self, port_name):
        """Connect to specified serial port"""
        try:
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.close()
            
            self.serial_port = serial.Serial(port_name, self.baud_rate, timeout=0.1)
            self.port_name = port_name
            self.status_message.emit(f"✓ Connected to {port_name}")
            self.connection_changed.emit(True)
            return True
        except Exception as e:
            self.status_message.emit(f"✗ Connection failed: {e}")
            self.connection_changed.emit(False)
            return False
    
    def disconnect(self):
        """Disconnect from serial port"""
        if self.serial_port and self.serial_port.is_open:
            self.serial_port.close()
        self.connection_changed.emit(False)
        self.status_message.emit("Disconnected")
    
    def run(self):
        """Main thread loop - reads and parses data"""
        self.running = True
        buffer = ""
        
        while self.running:
            if self.serial_port and self.serial_port.is_open:
                try:
                    if self.serial_port.in_waiting:
                        chunk = self.serial_port.read(self.serial_port.in_waiting).decode('utf-8', errors='ignore')
                        buffer += chunk
                        
                        while '\n' in buffer:
                            line, buffer = buffer.split('\n', 1)
                            line = line.strip()
                            
                            if not line:
                                continue
                            
                            if line.startswith('#'):
                                self.status_message.emit(line[1:].strip())
                            else:
                                try:
                                    parts = line.split(',')
                                    timestamp = float(parts[0]) / 1e6  # us to seconds
                                    samples = np.array([int(x) for x in parts[1:]], dtype=np.uint16)
                                    
                                    if len(samples) > 0:
                                        self.data_received.emit(samples, timestamp)
                                except (ValueError, IndexError):
                                    pass
                    
                    self.msleep(1)
                    
                except serial.SerialException:
                    self.status_message.emit("✗ Device disconnected")
                    self.connection_changed.emit(False)
                    if self.serial_port:
                        self.serial_port.close()
                        self.serial_port = None
                    self.msleep(1000)
            else:
                self.msleep(100)
    
    def stop(self):
        """Stop the thread"""
        self.running = False
        self.disconnect()


class SurfaceMicrophoneAnalyzer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Surface Microphone Analyzer")
        self.setGeometry(100, 100, 1600, 900)
        
        # Data storage - 30 seconds at 10kHz = 300,000 samples
        self.sample_rate = 10000  # Hz
        self.max_samples = 300000  # 30 seconds of data
        self.data_buffer = deque(maxlen=self.max_samples)
        self.time_buffer = deque(maxlen=self.max_samples)
        self.paused = False
        self.frozen_data = None
        self.frozen_time = None
        
        # Signal processing
        self.dc_offset = 2048
        self.highpass_enabled = True
        self.lowpass_enabled = False
        self.notch_enabled = True
        
        # Display options
        self.auto_scroll = True
        self.auto_zoom = True
        
        # Unified region selection for measurements and spike detection
        self.region_selection_mode = False
        self.selected_region = None  # (start_time, end_time)
        self.region_item = None
        self.detected_spikes = []
        self.spike_markers = []
        self._region_start = None  # Temporary storage during selection
        
        # Statistics
        self.total_samples_received = 0
        
        # Serial reader thread
        self.serial_reader = SerialReader()
        self.serial_reader.data_received.connect(self.on_data_received)
        self.serial_reader.status_message.connect(self.on_status_message)
        self.serial_reader.connection_changed.connect(self.on_connection_changed)
        self.serial_reader.start()
        
        self.setup_ui()
        
        # Update timer - 30 Hz
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.update_plot)
        self.update_timer.start(33)
        
        # Port scan timer
        self.port_scan_timer = QTimer()
        self.port_scan_timer.timeout.connect(self.scan_ports)
        self.port_scan_timer.start(2000)
        self.scan_ports()
    
    def setup_ui(self):
        """Create the user interface"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(5, 5, 5, 5)
        
        # Connection Controls
        conn_group = QGroupBox("Connection")
        conn_group.setMaximumHeight(70)
        conn_layout = QHBoxLayout()
        
        conn_layout.addWidget(QLabel("Port:"))
        self.port_combo = QComboBox()
        self.port_combo.setMinimumWidth(150)
        conn_layout.addWidget(self.port_combo)
        
        self.connect_btn = QPushButton("Connect")
        self.connect_btn.clicked.connect(self.toggle_connection)
        conn_layout.addWidget(self.connect_btn)
        
        self.status_label = QLabel("Not connected")
        self.status_label.setStyleSheet("color: #888;")
        conn_layout.addWidget(self.status_label)
        
        conn_layout.addStretch()
        conn_group.setLayout(conn_layout)
        main_layout.addWidget(conn_group)
        
        # Main Content Area with Splitter
        splitter = QSplitter(Qt.Horizontal)
        
        # Left side - Plot
        plot_container = QWidget()
        plot_layout = QVBoxLayout(plot_container)
        plot_layout.setContentsMargins(0, 0, 0, 0)

        self.plot_widget = pg.PlotWidget()
        self.plot_widget.useOpenGL(True)
        self.plot_widget.setBackground('w')
        self.plot_widget.setLabel('left', 'Amplitude', units='ADC')
        self.plot_widget.setLabel('bottom', 'Time', units='s')
        self.plot_widget.showGrid(x=True, y=True, alpha=0.3)
        self.plot_widget.setMouseEnabled(x=True, y=True)
        self.plot_widget.setClipToView(True)
        self.plot_curve = self.plot_widget.plot(
            pen=pg.mkPen(color='#2196F3', width=1),
            antialias=False
        )

        # Reference line at 0
        self.zero_line = pg.InfiniteLine(angle=0, movable=False, pen=pg.mkPen('#9E9E9E', width=1, style=Qt.DashLine))
        self.plot_widget.addItem(self.zero_line, ignoreBounds=True)
        
        # Crosshair for region selection
        self.vLine = pg.InfiniteLine(angle=90, movable=False, pen=pg.mkPen('#FF5722', width=1, style=Qt.DashLine))
        self.hLine = pg.InfiniteLine(angle=0, movable=False, pen=pg.mkPen('#FF5722', width=1, style=Qt.DashLine))
        self.plot_widget.addItem(self.vLine, ignoreBounds=True)
        self.plot_widget.addItem(self.hLine, ignoreBounds=True)
        self.vLine.setVisible(False)
        self.hLine.setVisible(False)
        
        # Mouse events
        self.plot_widget.scene().sigMouseMoved.connect(self.on_mouse_moved)
        self.plot_widget.scene().sigMouseClicked.connect(self.on_mouse_clicked)
        
        plot_layout.addWidget(self.plot_widget)
        splitter.addWidget(plot_container)
        
        # Right side - Controls
        self.control_panel = QWidget()
        self.control_panel.setMaximumWidth(350)
        control_panel_layout = QVBoxLayout(self.control_panel)
        
        self.collapse_btn = QPushButton("◀ Hide Controls")
        self.collapse_btn.clicked.connect(self.toggle_controls)
        self.collapse_btn.setMinimumHeight(30)
        control_panel_layout.addWidget(self.collapse_btn)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        
        # Control buttons
        control_group = QGroupBox("Controls")
        control_layout = QVBoxLayout()
        
        btn_row = QHBoxLayout()
        self.pause_btn = QPushButton("⏸ Pause")
        self.pause_btn.clicked.connect(self.toggle_pause)
        self.pause_btn.setEnabled(False)
        btn_row.addWidget(self.pause_btn)
        
        self.clear_btn = QPushButton("🗑 Clear")
        self.clear_btn.clicked.connect(self.clear_data)
        btn_row.addWidget(self.clear_btn)
        control_layout.addLayout(btn_row)
        
        self.region_checkbox = QCheckBox("Region Selection Mode")
        self.region_checkbox.stateChanged.connect(self.toggle_region_mode)
        control_layout.addWidget(self.region_checkbox)
        
        info_label = QLabel("Click twice to select a region for measurements and spike detection")
        info_label.setStyleSheet("color: #666; font-size: 10px;")
        info_label.setWordWrap(True)
        control_layout.addWidget(info_label)
        
        self.clear_region_btn = QPushButton("Clear Region")
        self.clear_region_btn.clicked.connect(self.clear_selected_region)
        self.clear_region_btn.setEnabled(False)
        control_layout.addWidget(self.clear_region_btn)
        
        control_group.setLayout(control_layout)
        scroll_layout.addWidget(control_group)
        
        # Spike Detection Group
        spike_group = QGroupBox("Spike Detection")
        spike_layout = QVBoxLayout()
        
        info_label2 = QLabel("Select region above, then adjust threshold and detect")
        info_label2.setStyleSheet("color: #666; font-size: 10px;")
        info_label2.setWordWrap(True)
        spike_layout.addWidget(info_label2)
        
        threshold_layout = QHBoxLayout()
        threshold_layout.addWidget(QLabel("Threshold:"))
        self.threshold_spin = QDoubleSpinBox()
        self.threshold_spin.setRange(1, 10000)
        self.threshold_spin.setValue(100.0)
        self.threshold_spin.setSuffix(" ADC")
        self.threshold_spin.setSingleStep(10)
        threshold_layout.addWidget(self.threshold_spin)
        spike_layout.addLayout(threshold_layout)
        
        min_distance_layout = QHBoxLayout()
        min_distance_layout.addWidget(QLabel("Min Distance:"))
        self.min_distance_spin = QDoubleSpinBox()
        self.min_distance_spin.setRange(0.001, 1.0)
        self.min_distance_spin.setValue(0.05)
        self.min_distance_spin.setSuffix(" s")
        self.min_distance_spin.setSingleStep(0.01)
        self.min_distance_spin.setDecimals(3)
        min_distance_layout.addWidget(self.min_distance_spin)
        spike_layout.addLayout(min_distance_layout)
        
        self.detect_btn = QPushButton("🔍 Detect Spikes")
        self.detect_btn.clicked.connect(self.detect_spikes)
        self.detect_btn.setEnabled(False)
        spike_layout.addWidget(self.detect_btn)
        
        self.spike_count_label = QLabel("No spikes detected")
        self.spike_count_label.setStyleSheet("color: #2196F3; font-weight: bold;")
        spike_layout.addWidget(self.spike_count_label)
        
        spike_group.setLayout(spike_layout)
        scroll_layout.addWidget(spike_group)
        
        # Signal Processing
        processing_group = QGroupBox("Signal Processing")
        processing_layout = QGridLayout()
        processing_layout.setSpacing(4)

        row = 0
        self.highpass_checkbox = QCheckBox("High-Pass")
        self.highpass_checkbox.setChecked(True)
        self.highpass_checkbox.stateChanged.connect(self.toggle_highpass)
        processing_layout.addWidget(self.highpass_checkbox, row, 0)

        self.highpass_spin = QDoubleSpinBox()
        self.highpass_spin.setRange(0.1, 1000)
        self.highpass_spin.setValue(10.0)
        self.highpass_spin.setSuffix(" Hz")
        self.highpass_spin.setSingleStep(0.5)
        self.highpass_spin.setMaximumWidth(100)
        processing_layout.addWidget(self.highpass_spin, row, 1)

        row += 1
        self.lowpass_checkbox = QCheckBox("Low-Pass")
        self.lowpass_checkbox.setChecked(False)
        self.lowpass_checkbox.stateChanged.connect(self.toggle_lowpass)
        processing_layout.addWidget(self.lowpass_checkbox, row, 0)

        self.lowpass_spin = QDoubleSpinBox()
        self.lowpass_spin.setRange(1, 5000)
        self.lowpass_spin.setValue(500.0)
        self.lowpass_spin.setSuffix(" Hz")
        self.lowpass_spin.setSingleStep(10)
        processing_layout.addWidget(self.lowpass_spin, row, 1)

        row += 1
        self.notch_checkbox = QCheckBox("Notch Filter")
        self.notch_checkbox.setChecked(True)
        self.notch_checkbox.stateChanged.connect(self.toggle_notch)
        processing_layout.addWidget(self.notch_checkbox, row, 0, 1, 2)

        row += 1
        processing_layout.addWidget(QLabel("Freq:"), row, 0)
        self.notch_freq_spin = QDoubleSpinBox()
        self.notch_freq_spin.setRange(1, 5000)
        self.notch_freq_spin.setValue(50.0)
        self.notch_freq_spin.setSuffix(" Hz")
        processing_layout.addWidget(self.notch_freq_spin, row, 1)

        row += 1
        processing_layout.addWidget(QLabel("Q:"), row, 0)
        self.notch_q_spin = QDoubleSpinBox()
        self.notch_q_spin.setRange(1, 100)
        self.notch_q_spin.setValue(30.0)
        self.notch_q_spin.setSingleStep(1)
        processing_layout.addWidget(self.notch_q_spin, row, 1)

        processing_group.setLayout(processing_layout)
        scroll_layout.addWidget(processing_group)

        # Display settings
        display_group = QGroupBox("Display")
        display_layout = QVBoxLayout()
        
        h_layout = QHBoxLayout()
        h_layout.addWidget(QLabel("Window:"))
        self.window_spin = QDoubleSpinBox()
        self.window_spin.setRange(0.1, 30.0)
        self.window_spin.setValue(5.0)
        self.window_spin.setSuffix(" s")
        self.window_spin.setSingleStep(0.5)
        h_layout.addWidget(self.window_spin)
        display_layout.addLayout(h_layout)
        
        self.auto_scroll_checkbox = QCheckBox("Auto-Scroll")
        self.auto_scroll_checkbox.setChecked(True)
        self.auto_scroll_checkbox.stateChanged.connect(self.toggle_auto_scroll)
        display_layout.addWidget(self.auto_scroll_checkbox)
        
        self.auto_zoom_checkbox = QCheckBox("Auto Y-Zoom")
        self.auto_zoom_checkbox.setChecked(True)
        self.auto_zoom_checkbox.stateChanged.connect(self.toggle_auto_zoom)
        display_layout.addWidget(self.auto_zoom_checkbox)
        
        display_group.setLayout(display_layout)
        scroll_layout.addWidget(display_group)
        
        # Statistics
        stats_group = QGroupBox("Statistics")
        stats_layout = QGridLayout()
        stats_layout.setSpacing(4)

        self.stats_labels = {}
        stat_names = [
            ('samples', 'Samples:'),
            ('duration', 'Duration:'),
            ('rate', 'Rate:'),
            ('rms', 'RMS:'),
            ('peak_to_peak', 'Peak-Peak:')
        ]

        for i, (key, label) in enumerate(stat_names):
            label_widget = QLabel(label)
            label_widget.setStyleSheet("font-weight: bold;")
            stats_layout.addWidget(label_widget, i, 0)
            
            value_label = QLabel("---")
            value_label.setFont(QFont("Courier", 9))
            stats_layout.addWidget(value_label, i, 1)
            self.stats_labels[key] = value_label

        stats_group.setLayout(stats_layout)
        scroll_layout.addWidget(stats_group)
        
        # Measurement Results
        measurement_group = QGroupBox("Region Measurements")
        measurement_layout = QVBoxLayout()
        
        self.measurement_labels = {}
        measurement_info = [
            ('region', 'Region:', 'No region'),
            ('duration', 'Duration:', '---'),
            ('samples', 'Samples:', '---'),
            ('rms', 'RMS:', '---'),
            ('peak_to_peak', 'Peak-Peak:', '---'),
            ('avg_freq', 'Avg Freq:', '---')
        ]
        
        for key, label, default in measurement_info:
            h_layout = QHBoxLayout()
            label_widget = QLabel(label)
            label_widget.setStyleSheet("font-weight: bold;")
            label_widget.setMinimumWidth(70)
            h_layout.addWidget(label_widget)
            
            value_label = QLabel(default)
            value_label.setFont(QFont("Courier", 8))
            value_label.setWordWrap(True)
            h_layout.addWidget(value_label)
            
            measurement_layout.addLayout(h_layout)
            self.measurement_labels[key] = value_label
        
        measurement_group.setLayout(measurement_layout)
        scroll_layout.addWidget(measurement_group)
        
        scroll_layout.addStretch()
        scroll.setWidget(scroll_content)
        control_panel_layout.addWidget(scroll)
        
        splitter.addWidget(self.control_panel)
        splitter.setSizes([1280, 320])
        
        main_layout.addWidget(splitter)
    
    def toggle_controls(self):
        """Toggle visibility of control panel"""
        scroll_widget = self.control_panel.findChild(QScrollArea)
        
        if scroll_widget and scroll_widget.isVisible():
            scroll_widget.hide()
            self.collapse_btn.setText("▶")
            self.control_panel.setMaximumWidth(40)
        else:
            if scroll_widget:
                scroll_widget.show()
            self.collapse_btn.setText("◀ Hide Controls")
            self.control_panel.setMaximumWidth(16777215)
    
    def scan_ports(self):
        """Scan for available serial ports"""
        ports = serial.tools.list_ports.comports()
        current_ports = [self.port_combo.itemText(i) for i in range(self.port_combo.count())]
        available_ports = [port.device for port in ports]
        
        if set(current_ports) != set(available_ports):
            current_selection = self.port_combo.currentText()
            self.port_combo.clear()
            self.port_combo.addItems(available_ports)
            
            if current_selection in available_ports:
                self.port_combo.setCurrentText(current_selection)
    
    def toggle_connection(self):
        """Connect or disconnect from serial port"""
        if self.connect_btn.text() == "Connect":
            port = self.port_combo.currentText()
            if port and self.serial_reader.connect_to_port(port):
                self.connect_btn.setText("Disconnect")
                self.pause_btn.setEnabled(True)
        else:
            self.serial_reader.disconnect()
            self.connect_btn.setText("Connect")
            self.pause_btn.setEnabled(False)
    
    def on_connection_changed(self, connected):
        """Handle connection status changes"""
        if not connected and self.connect_btn.text() == "Disconnect":
            self.connect_btn.setText("Connect")
            self.pause_btn.setEnabled(False)
    
    def on_data_received(self, samples, timestamp):
        """Handle incoming data from serial port"""
        if not self.paused:
            self.total_samples_received += len(samples)
            
            dt = 1.0 / self.sample_rate
            times = timestamp + np.arange(len(samples)) * dt
            
            self.data_buffer.extend(samples)
            self.time_buffer.extend(times)
    
    def on_status_message(self, message):
        """Display status message"""
        self.status_label.setText(message)
        if "✓" in message:
            self.status_label.setStyleSheet("color: #4CAF50; font-weight: bold;")
        elif "✗" in message:
            self.status_label.setStyleSheet("color: #F44336; font-weight: bold;")
        else:
            self.status_label.setStyleSheet("color: #2196F3;")
    
    def toggle_pause(self):
        """Pause or resume data acquisition"""
        self.paused = not self.paused
        if self.paused:
            self.pause_btn.setText("▶ Resume")
            self.frozen_data = np.array(self.data_buffer)
            self.frozen_time = np.array(self.time_buffer)
        else:
            self.pause_btn.setText("⏸ Pause")
            self.frozen_data = None
            self.frozen_time = None
            if self.auto_scroll:
                self.plot_widget.enableAutoRange(axis='x')
            if self.auto_zoom:
                self.plot_widget.enableAutoRange(axis='y')
    
    def clear_data(self):
        """Clear all data buffers"""
        self.data_buffer.clear()
        self.time_buffer.clear()
        self.frozen_data = None
        self.frozen_time = None
        self.total_samples_received = 0
        self.clear_selected_region()
    
    def toggle_region_mode(self, state):
        """Enable/disable region selection mode"""
        self.region_selection_mode = (state == 2)
        if not self.region_selection_mode:
            self._region_start = None
            self.vLine.setVisible(False)
            self.hLine.setVisible(False)
    
    def toggle_highpass(self, state):
        self.highpass_enabled = (state == 2)
    
    def toggle_lowpass(self, state):
        self.lowpass_enabled = (state == 2)
    
    def toggle_notch(self, state):
        self.notch_enabled = (state == 2)
    
    def toggle_auto_scroll(self, state):
        self.auto_scroll = (state == 2)
    
    def toggle_auto_zoom(self, state):
        self.auto_zoom = (state == 2)
    
    def clear_selected_region(self):
        """Clear selected region and all associated data"""
        self.selected_region = None
        self._region_start = None
        self.detected_spikes = []
        
        # Remove region visualization
        if self.region_item and self.region_item in self.plot_widget.items():
            self.plot_widget.removeItem(self.region_item)
        self.region_item = None
        
        # Remove spike markers
        for marker in self.spike_markers:
            if marker in self.plot_widget.items():
                self.plot_widget.removeItem(marker)
        self.spike_markers = []
        
        # Reset labels
        self.spike_count_label.setText("No spikes detected")
        self.measurement_labels['region'].setText('No region')
        self.measurement_labels['duration'].setText('---')
        self.measurement_labels['samples'].setText('---')
        self.measurement_labels['rms'].setText('---')
        self.measurement_labels['peak_to_peak'].setText('---')
        self.measurement_labels['avg_freq'].setText('---')
        
        self.detect_btn.setEnabled(False)
        self.clear_region_btn.setEnabled(False)
    
    def detect_spikes(self):
        """Detect spikes in the selected region - FIXED: No more duplicates"""
        if not self.selected_region:
            QMessageBox.warning(self, "No Region", "Please select a region first.")
            return
        
        # Get data
        if self.paused and self.frozen_data is not None:
            data = self.frozen_data
            times = self.frozen_time
        else:
            if len(self.data_buffer) == 0:
                QMessageBox.warning(self, "No Data", "No data available.")
                return
            data = np.array(self.data_buffer)
            times = np.array(self.time_buffer)
        
        # Process signal
        processed_data = self.process_signal(data)
        
        # Extract region
        start_time, end_time = self.selected_region
        mask = (times >= start_time) & (times <= end_time)
        
        if not np.any(mask):
            QMessageBox.warning(self, "No Data", "No data in selected region.")
            return
        
        region_times = times[mask]
        region_data = processed_data[mask]
        
        # FIXED: Detect spikes without duplicates
        # Find where absolute value exceeds threshold
        threshold = self.threshold_spin.value()
        min_distance_sec = self.min_distance_spin.value()
        min_distance_samples = int(min_distance_sec * self.sample_rate)
        
        # Work with absolute values to find all spikes regardless of direction
        abs_data = np.abs(region_data)
        
        # Find peaks in absolute value
        peaks, _ = scipy_signal.find_peaks(abs_data, height=threshold, distance=min_distance_samples)
        
        if len(peaks) == 0:
            QMessageBox.information(self, "No Spikes", f"No spikes found above threshold {threshold} ADC.")
            return
        
        # Clear previous markers
        for marker in self.spike_markers:
            if marker in self.plot_widget.items():
                self.plot_widget.removeItem(marker)
        self.spike_markers = []
        
        # Analyze each spike
        self.detected_spikes = []
        
        for peak_idx in peaks:
            peak_time = region_times[peak_idx]
            
            # Find local min and max around this peak
            window = min_distance_samples // 2
            start_idx = max(0, peak_idx - window)
            end_idx = min(len(region_data), peak_idx + window + 1)
            
            local_data = region_data[start_idx:end_idx]
            
            if len(local_data) > 0:
                local_max = np.max(local_data)
                local_min = np.min(local_data)
                
                self.detected_spikes.append({
                    'time': peak_time,
                    'max': local_max,
                    'min': local_min,
                    'amplitude': local_max - local_min
                })
                
                # Add visual marker
                marker = pg.InfiniteLine(
                    angle=90, 
                    pos=peak_time,
                    pen=pg.mkPen('#4CAF50', width=2, style=Qt.DashLine)
                )
                self.plot_widget.addItem(marker)
                self.spike_markers.append(marker)
        
        # Update label
        self.spike_count_label.setText(f"✓ Found {len(self.detected_spikes)} spikes")
        
        # Show results dialog
        dialog = SpikeResultsDialog(self.detected_spikes, self)
        dialog.exec_()
    
    def on_mouse_clicked(self, event):
        """Handle mouse clicks for region selection"""
        if not self.region_selection_mode:
            return
        
        if event.button() != Qt.LeftButton:
            return
        
        # Get data
        if self.paused and self.frozen_data is not None:
            times = self.frozen_time
        else:
            if len(self.data_buffer) == 0:
                return
            times = np.array(self.time_buffer)
        
        if len(times) == 0:
            return
        
        mouse_point = self.plot_widget.plotItem.vb.mapSceneToView(event.scenePos())
        click_time = mouse_point.x()
        
        if self._region_start is None:
            # First click - start region
            self._region_start = click_time
            
            # Show temporary marker
            if self.region_item:
                self.plot_widget.removeItem(self.region_item)
            
            self.region_item = pg.InfiniteLine(
                angle=90, 
                pos=click_time,
                pen=pg.mkPen('#FF9800', width=2)
            )
            self.plot_widget.addItem(self.region_item)
        else:
            # Second click - end region
            region_end = click_time
            start = min(self._region_start, region_end)
            end = max(self._region_start, region_end)
            
            self.selected_region = (start, end)
            self._region_start = None
            
            # Remove old line and show shaded region
            if self.region_item:
                self.plot_widget.removeItem(self.region_item)
            
            self.region_item = pg.LinearRegionItem(
                values=[start, end],
                brush=pg.mkBrush(255, 152, 0, 50),
                movable=False
            )
            self.plot_widget.addItem(self.region_item)
            
            self.detect_btn.setEnabled(True)
            self.clear_region_btn.setEnabled(True)
            
            # Update region measurements
            self.update_region_measurements()
    
    def update_region_measurements(self):
        """Update measurement statistics for the selected region"""
        if not self.selected_region:
            return
        
        # Get data
        if self.paused and self.frozen_data is not None:
            data = self.frozen_data
            times = self.frozen_time
        else:
            if len(self.data_buffer) == 0:
                return
            data = np.array(self.data_buffer)
            times = np.array(self.time_buffer)
        
        # Process signal
        processed_data = self.process_signal(data)
        
        # Extract region
        start_time, end_time = self.selected_region
        mask = (times >= start_time) & (times <= end_time)
        
        if not np.any(mask):
            return
        
        region_times = times[mask]
        region_data = processed_data[mask]
        
        # Calculate statistics
        duration = end_time - start_time
        num_samples = len(region_data)
        rms = np.sqrt(np.mean(region_data**2))
        p2p = np.max(region_data) - np.min(region_data)
        
        # Try to estimate frequency using zero crossings
        zero_crossings = np.where(np.diff(np.sign(region_data)))[0]
        if len(zero_crossings) > 1:
            avg_period = np.mean(np.diff(region_times[zero_crossings])) * 2  # *2 for full period
            if avg_period > 0:
                avg_freq = 1.0 / avg_period
                freq_text = f"{avg_freq:.2f} Hz"
            else:
                freq_text = "---"
        else:
            freq_text = "---"
        
        # Update labels
        self.measurement_labels['region'].setText(f"{start_time:.3f} to {end_time:.3f}s")
        self.measurement_labels['duration'].setText(f"{duration:.4f} s")
        self.measurement_labels['samples'].setText(f"{num_samples:,}")
        self.measurement_labels['rms'].setText(f"{rms:.1f} ADC")
        self.measurement_labels['peak_to_peak'].setText(f"{p2p:.1f} ADC")
        self.measurement_labels['avg_freq'].setText(freq_text)
    
    def on_mouse_moved(self, pos):
        """Handle mouse movement for crosshair"""
        if self.region_selection_mode and not self.paused:
            mouse_point = self.plot_widget.plotItem.vb.mapSceneToView(pos)
            self.vLine.setPos(mouse_point.x())
            self.hLine.setPos(mouse_point.y())
            self.vLine.setVisible(True)
            self.hLine.setVisible(True)
        else:
            self.vLine.setVisible(False)
            self.hLine.setVisible(False)
    
    def process_signal(self, data):
        """Apply signal processing"""
        processed = data.copy().astype(float)
        
        any_filter_active = (self.highpass_enabled or self.lowpass_enabled or self.notch_enabled)
        
        if any_filter_active:
            self.dc_offset = np.mean(processed)
            processed = processed - self.dc_offset
        else:
            self.dc_offset = 2048
        
        if self.highpass_enabled and len(processed) > 100:
            cutoff = self.highpass_spin.value()
            sos = scipy_signal.butter(4, cutoff, 'hp', fs=self.sample_rate, output='sos')
            processed = scipy_signal.sosfiltfilt(sos, processed)
        
        if self.lowpass_enabled and len(processed) > 100:
            cutoff = self.lowpass_spin.value()
            sos = scipy_signal.butter(4, cutoff, 'lp', fs=self.sample_rate, output='sos')
            processed = scipy_signal.sosfiltfilt(sos, processed)
        
        if self.notch_enabled and len(processed) > 100:
            freq = self.notch_freq_spin.value()
            q = self.notch_q_spin.value()
            b, a = scipy_signal.iirnotch(freq, q, self.sample_rate)
            processed = scipy_signal.filtfilt(b, a, processed)
        
        return processed
    
    def update_plot(self):
        """Update the plot and statistics"""
        if self.paused and self.frozen_data is not None:
            data = self.frozen_data.astype(float)
            times = self.frozen_time
        else:
            if len(self.data_buffer) == 0:
                return
            data = np.array(self.data_buffer).astype(float)
            times = np.array(self.time_buffer)
        
        window_size = self.window_spin.value()
        if len(times) > 0:
            cutoff_time = times[-1] - window_size
            mask = times >= cutoff_time
            data = data[mask]
            times = times[mask]
        
        if len(data) == 0:
            return
        
        # Process signal
        processed_data = self.process_signal(data)
        
        # Downsample for display if too many points (performance boost)
        if len(processed_data) > 10000:
            step = max(1, len(processed_data) // 10000)
            display_times = times[::step]
            display_data = processed_data[::step]
        else:
            display_times = times
            display_data = processed_data
        
        # Update main plot
        self.plot_curve.setData(display_times, display_data)
        
        # Handle auto-scroll
        if self.auto_scroll and not self.paused:
            if len(times) > 0:
                self.plot_widget.setXRange(times[-1] - window_size, times[-1], padding=0)
        
        # Handle auto-zoom
        if self.auto_zoom:
            if len(processed_data) > 0:
                data_min = np.min(processed_data)
                data_max = np.max(processed_data)
                padding = (data_max - data_min) * 0.1
                self.plot_widget.setYRange(data_min - padding, data_max + padding, padding=0)
        
        # Update zero reference line
        self.zero_line.setPos(0 if any([self.highpass_enabled, self.lowpass_enabled, 
                                        self.notch_enabled]) else self.dc_offset)
        
        # Update statistics
        self.stats_labels['samples'].setText(f"{self.total_samples_received:,}")
        
        # Duration
        if len(self.time_buffer) > 1:
            duration = self.time_buffer[-1] - self.time_buffer[0]
            self.stats_labels['duration'].setText(f"{duration:.2f} s")
        else:
            self.stats_labels['duration'].setText("---")
        
        # Actual sample rate
        if len(self.data_buffer) > 0 and len(self.time_buffer) > 1:
            time_span = self.time_buffer[-1] - self.time_buffer[0]
            if time_span > 0:
                actual_rate = len(self.data_buffer) / time_span
                self.stats_labels['rate'].setText(f"{actual_rate:.0f} sps")
            else:
                self.stats_labels['rate'].setText("--- sps")
        else:
            self.stats_labels['rate'].setText("--- sps")
        
        # Peak-to-peak
        if len(processed_data) > 0:
            p2p = np.max(processed_data) - np.min(processed_data)
            self.stats_labels['peak_to_peak'].setText(f"{p2p:.1f} ADC")
        else:
            self.stats_labels['peak_to_peak'].setText("---")
        
        # RMS (Root Mean Square) - signal energy
        if len(processed_data) > 0:
            rms = np.sqrt(np.mean(processed_data**2))
            self.stats_labels['rms'].setText(f"{rms:.1f} ADC")
        else:
            self.stats_labels['rms'].setText("---")
    
    def closeEvent(self, event):
        """Clean up on window close"""
        self.serial_reader.stop()
        self.serial_reader.wait()
        event.accept()


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = SurfaceMicrophoneAnalyzer()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
