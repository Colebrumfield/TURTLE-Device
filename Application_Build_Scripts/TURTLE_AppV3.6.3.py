from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QComboBox,
    QCheckBox, QLineEdit, QMessageBox, QFileDialog, QGridLayout, QDialog, QRadioButton, QButtonGroup, QGroupBox
)
from PyQt6.QtGui import QFont, QIcon
from PyQt6.QtCore import QTimer, Qt

import os
import sys
from time import time as current_time
from datetime import datetime

import pandas as pd
import matplotlib.pyplot as plt

import serial
from serial.tools import list_ports
from threading import Thread

class ThermocoupleUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AGGRC TURTLE App Version 3.6.3")
        self.setGeometry(100, 100, 400, 300)
        self.setCentralWidget(QWidget())
        self.layout = QVBoxLayout()
        self.centralWidget().setLayout(self.layout)

        #icon
        self.resource_dir = os.path.join(os.path.dirname(__file__), 'resources')
        main_window_icon_path = os.path.join(self.resource_dir, 'Turtle_Icon.png')
        self.setWindowIcon(QIcon(main_window_icon_path))

        self.default_connections()
        self.setup_ui()
        self.start_reading_data()

    def default_connections(self):
        #arduino connection
        self.connected = False
        self.ser = None
        self.connect_to_arduino()

        #timing
        self.start_time = None
        self.is_recording = False

        #list for data
        self.temp_data = []
        self.temps_f = []
        self.temps_c = []
        self.connection_status_labels = []

        #timing
        self.timestamp_counter = 0
        self.sampling_rate = 1
        
        #for enabling start button
        self.first_connection = True

        #settings window defaults
        self.temp_unit = "C"
        self.theme = 'light'
        self.setStyleSheet('')

    def setup_ui(self):
        self.font = QFont("Arial", 16)

        # Layouts
        main_layout = QHBoxLayout()
        self.layout.addLayout(main_layout)
        
        # Thermocouple Frames
        tc_frame = QWidget()
        tc_layout = QGridLayout()
        tc_frame.setLayout(tc_layout)
        main_layout.addWidget(tc_frame)

        for i in range(2):
            row = i // 2
            column = i % 2
            
            # Create a QGroupBox to hold the connection status and temperature labels
            group_box = QGroupBox(f"Thermocouple {i+1}:")
            group_box.setFont(self.font)
            group_layout = QVBoxLayout()
            group_box.setLayout(group_layout)
            tc_layout.addWidget(group_box, row, column)

            connection_status_label = QLabel("Not Connected")
            connection_status_label.setFont(self.font)
            connection_status_label.setStyleSheet("background-color: #333; color: #808080; padding: 5px; border-radius: 10px;")
            group_layout.addWidget(connection_status_label)
            self.connection_status_labels.append(connection_status_label)

            temp_c = QLabel("N/A °C", font=self.font)
            group_layout.addWidget(temp_c)
            self.temps_c.append(temp_c)

        # Sampling Rate ComboBox (under the left thermocouple displays)
        self.sampling_rate_combobox = QComboBox()
        sampling_rate_options = [f"Sample every {i} second(s)" for i in range(1, 6)]
        sampling_rate_options.append("Max (≈ 3-4 samples/sec)")
        self.sampling_rate_combobox.addItems(sampling_rate_options)
        self.sampling_rate_combobox.setCurrentText("Sample every 1 second(s)")
        self.sampling_rate_combobox.currentTextChanged.connect(self.update_sampling_rate)
        self.sampling_rate_combobox.setFont(self.font)
        self.sampling_rate_combobox.setFixedWidth(300)
        tc_layout.addWidget(self.sampling_rate_combobox, 2, 0, 1, 1)  # Place in row 2, column 0

        # Thermocouple Type ComboBox (under the right thermocouple displays)
        self.thermocouple_type_combobox = QComboBox()
        thermocouple_type_options = ["Thermocouple Type " + type for type in ["K", "J", "T", "E", "N", "S", "R", "B"]]
        self.thermocouple_type_combobox.addItems(thermocouple_type_options)
        self.thermocouple_type_combobox.setCurrentText("Thermocouple Type T")
        self.thermocouple_type_combobox.currentTextChanged.connect(self.update_tc_type)
        self.thermocouple_type_combobox.setFont(self.font)
        self.thermocouple_type_combobox.setFixedWidth(300)
        tc_layout.addWidget(self.thermocouple_type_combobox, 2, 1, 1, 1)  # Place in row 2, column 1
     
        # Options Frame
        options_frame = QWidget()
        options_layout = QVBoxLayout()
        options_frame.setLayout(options_layout)
        main_layout.addWidget(options_frame)

        # Start Recording Button
        self.start_button = QPushButton("Start Recording")
        self.start_button.setFont(self.font)
        self.start_button.setStyleSheet("background-color: green;")
        self.start_button.clicked.connect(self.toggle_recording)
        options_layout.addWidget(self.start_button)

        self.start_button.setEnabled(False)

        #elapsed time
        elapsed_time_box = QGroupBox() #Create a QGroupBox to contain the elapsed time label
        elapsed_time_layout = QVBoxLayout() # Set a layout for the QGroupBox
        elapsed_time_box.setLayout(elapsed_time_layout)
        self.elapsed_time_label = QLabel("Time Elapsed: 0s", font=self.font) # Create and add the elapsed time label to the layout
        elapsed_time_layout.addWidget(self.elapsed_time_label)
        options_layout.addWidget(elapsed_time_box) # Add the QGroupBox to the options layout

        # Graph Button
        self.graph_button = QPushButton("Graph Last Data Points")
        self.graph_button.setFont(self.font)
        self.graph_button.setStyleSheet("background-color: purple;")
        graph_icon_path = os.path.join(self.resource_dir, 'Graph_Icon.ico')
        self.graph_button.setIcon(QIcon(graph_icon_path))  # Set the icon for the button
        self.graph_button.clicked.connect(self.show_graph)
        options_layout.addWidget(self.graph_button)

        # Export Button
        self.export_button = QPushButton("Export Last Data to Excel")
        self.export_button.setFont(self.font)
        self.export_button.setStyleSheet("background-color: blue;")
        excel_icon_path = os.path.join(self.resource_dir, 'Excel_Icon.ico')
        self.export_button.setIcon(QIcon(excel_icon_path))  # Set the icon for the button
        self.export_button.clicked.connect(self.export_to_excel)
        options_layout.addWidget(self.export_button)

        # Cooling Rate Options Box
        cooling_rate_box = QGroupBox()  # "Cooling Rate Options:"
        cooling_rate_layout = QVBoxLayout()
        cooling_rate_box.setLayout(cooling_rate_layout)

        # First Cooling Rate Calculation
        self.calculate_cooling_check_1 = QCheckBox("Calculate Cooling Rate (Thermocouple 1)")
        self.calculate_cooling_check_1.setFont(self.font)
        cooling_rate_layout.addWidget(self.calculate_cooling_check_1)

        self.entry_interval_1 = QLineEdit()
        self.entry_interval_1.setPlaceholderText("Starting Temp 1 (°C)")
        self.entry_interval_1.setFont(self.font)
        cooling_rate_layout.addWidget(self.entry_interval_1)

        self.entry_interval_2 = QLineEdit()
        self.entry_interval_2.setPlaceholderText("Ending Temp 1 (°C)")
        self.entry_interval_2.setFont(self.font)
        cooling_rate_layout.addWidget(self.entry_interval_2)

        # Second Cooling Rate Calculation
        self.calculate_cooling_check_2 = QCheckBox("Calculate Cooling Rate (Thermocouple 2)")
        self.calculate_cooling_check_2.setFont(self.font)
        cooling_rate_layout.addWidget(self.calculate_cooling_check_2)

        self.entry_interval_3 = QLineEdit()
        self.entry_interval_3.setPlaceholderText("Starting Temp 2 (°C)")
        self.entry_interval_3.setFont(self.font)
        cooling_rate_layout.addWidget(self.entry_interval_3)

        self.entry_interval_4 = QLineEdit()
        self.entry_interval_4.setPlaceholderText("Ending Temp 2 (°C)")
        self.entry_interval_4.setFont(self.font)
        cooling_rate_layout.addWidget(self.entry_interval_4)

        # Add Cooling Rate Box to Options Layout
        options_layout.addWidget(cooling_rate_box)

    def update_sampling_rate(self, rate):
        if "Max" in rate:
            rate_value = 0
        else:
            rate_value = int(rate.split(" ")[2])
        self.sampling_rate = rate_value
        self.send_to_arduino(f"RATE:{rate_value};")

    def update_tc_type(self, type):
        tc_type = type.split(" ")[2]
        self.send_to_arduino(f"TYPE:{tc_type}")

    def start_reading_data(self):
        self.reading_thread = Thread(target=self.read_data, daemon=True)
        self.reading_thread.start()

    def read_data(self):
        while True:
            if self.connected:
                try:
                    if self.ser.in_waiting > 0:
                        line = self.ser.readline().decode('utf-8').strip()
                        if "STATUS:" in line:
                            _, statuses = line.split("STATUS:")
                            temp_data = statuses.split(",")[:-1]
                            data_collected = []
                            for data in temp_data:
                                tc, status = data.split(":")
                                index = int(tc[1]) - 1

                                if status == "Not Connected":
                                    self.connection_status_labels[index].setText("Not Connected")
                                    self.connection_status_labels[index].setStyleSheet("background-color: #333; color:rgb(203, 199, 199); padding: 5px; border-radius: 10px;")
                                    self.temps_c[index].setText("N/A °C")
                                else:
                                    if self.first_connection:
                                        self.start_button.setEnabled(True)
                                        self.first_connection = False
                                    self.connection_status_labels[index].setText("Connected")
                                    self.connection_status_labels[index].setStyleSheet("background-color: green;")
                                    temp_c = float(status)
                                    self.temps_c[index].setText(f"{temp_c:.2f}°C")

                                    if self.is_recording:
                                        elapsed_time = current_time() - self.start_time
                                        new_reading = {
                                            'timestamp': round(elapsed_time, 2),
                                            'tc_id': index + 1,
                                            'temp_c': temp_c
                                        }
                                        self.temp_data.append(new_reading)
                                        self.update_graph()
                except Exception as e:
                    pass

    def toggle_recording(self):
        if self.is_recording:
            self.is_recording = False
            self.start_button.setText("Start Recording")
            self.start_button.setStyleSheet("background-color: green;")
        else:
            self.is_recording = True
            self.start_button.setText("Stop Recording")
            self.start_button.setStyleSheet("background-color: red;")
            self.start_time = current_time()
            self.temp_data = []
            self.update_elapsed_time()

    def update_elapsed_time(self):
        if self.is_recording:
            elapsed_time = round(current_time() - self.start_time, 2)
            self.elapsed_time_label.setText(f"Time Elapsed: {elapsed_time}s")
            QTimer.singleShot(1000, self.update_elapsed_time)

    def calculate_cooling_rate_tc1(self, df, interval1, interval2):
        return self._calculate_cooling_rate(df, interval1, interval2, "Thermocouple 1")

    def calculate_cooling_rate_tc2(self, df, interval3, interval4):
        return self._calculate_cooling_rate(df, interval3, interval4, "Thermocouple 2")

    def _calculate_cooling_rate(self, df, interval_start, interval_end, label):
        try:
            interval_start = float(interval_start)
            interval_end = float(interval_end)
        except ValueError:
            QMessageBox.information(self, "Error", f"Please enter valid numerical values for {label}.")
            return None

        cooling_rates = {}
        grouped = df.groupby('tc_id')

        for tc_id, group in grouped:
            group = group.reset_index(drop=True)

            # Find the closest temperatures to the specified intervals
            interval1_index = (group['temp_c'] - interval_start).abs().idxmin()
            interval2_index = (group['temp_c'] - interval_end).abs().idxmin()

            interval1_row = group.loc[interval1_index]
            interval2_row = group.loc[interval2_index]

            if not interval1_row.empty and not interval2_row.empty:
                interval1_temp = interval1_row['temp_c']
                interval2_temp = interval2_row['temp_c']
                interval1_time = interval1_row['timestamp']
                interval2_time = interval2_row['timestamp']
                cooling_rate = ((interval2_temp - interval1_temp) / (interval2_time - interval1_time)) * 60
                cooling_rate = round(cooling_rate, 2)  # Limit to 2 decimal places

                cooling_rates[tc_id] = {
                    'cooling_rate': cooling_rate,
                    'interval1_temp': interval1_temp,
                    'interval2_temp': interval2_temp,
                    'interval1_time': interval1_time,
                    'interval2_time': interval2_time
                }
            else:
                cooling_rates[tc_id] = None

        if not cooling_rates:
            QMessageBox.information(self, "Error", f"No data found for {label} intervals.")
            return None

        return cooling_rates

    def show_graph(self):
        if not self.temp_data:
            QMessageBox.information(self, "No Data", "No temperature data to plot.")
            return

        df = pd.DataFrame(self.temp_data)

        plt.figure(figsize=(10, 6))

        cooling_rates_tc1 = {}
        cooling_rates_tc2 = {}

        # Calculate cooling rate for Thermocouple 1 if checked
        if self.calculate_cooling_check_1.isChecked():
            interval1 = self.entry_interval_1.text()
            interval2 = self.entry_interval_2.text()
            cooling_rates_tc1 = self.calculate_cooling_rate_tc1(df, interval1, interval2)

        # Calculate cooling rate for Thermocouple 2 if checked
        if self.calculate_cooling_check_2.isChecked():
            interval3 = self.entry_interval_3.text()
            interval4 = self.entry_interval_4.text()
            cooling_rates_tc2 = self.calculate_cooling_rate_tc2(df, interval3, interval4)

        # Plot the temperature data for each thermocouple
        for tc_id in df['tc_id'].unique():
            tc_df = df[df['tc_id'] == tc_id]
            plt.plot(tc_df['timestamp'], tc_df['temp_c'], label=f'Thermocouple {tc_id}')

        # Add legend
        plt.legend(loc='upper left')

        # Determine the position for the cooling rate text boxes
        xlims = plt.xlim()  # Get x-axis limits
        ylims = plt.ylim()  # Get y-axis limits

        # Place the cooling rate text boxes in the upper right corner of the plot
        text_x = xlims[1] - 0.1 * (xlims[1] - xlims[0])  # Near the right edge
        text_y = ylims[1] - 0.2 * (ylims[1] - ylims[0])  # Near the top, but not too close

        # Display cooling rate only for Thermocouple 1 if checked
        if self.calculate_cooling_check_1.isChecked():
            if 1 in cooling_rates_tc1:  # Only display if TC1 has a cooling rate
                data = cooling_rates_tc1[1]
                if data:
                    cooling_rate = data['cooling_rate']
                    interval1_temp = data['interval1_temp']
                    interval2_temp = data['interval2_temp']

                    textstr = f'TC 1 Cooling rate: {cooling_rate:.2f}°{self.temp_unit}/min\n' \
                            f'Interval: {interval1_temp:.2f}°{self.temp_unit} to {interval2_temp:.2f}°{self.temp_unit}'

                    plt.text(text_x, text_y, textstr, fontsize=10, color='red',
                            verticalalignment='top', horizontalalignment='right',
                            bbox=dict(facecolor='white', alpha=0.5))
                    text_y -= 0.1 * (ylims[1] - ylims[0])  # Adjust vertical spacing

        # Display cooling rate only for Thermocouple 2 if checked
        if self.calculate_cooling_check_2.isChecked():
            if 2 in cooling_rates_tc2:  # Only display if TC2 has a cooling rate
                data = cooling_rates_tc2[2]
                if data:
                    cooling_rate = data['cooling_rate']
                    interval3_temp = data['interval1_temp']
                    interval4_temp = data['interval2_temp']

                    textstr = f'TC 2 Cooling rate: {cooling_rate:.2f}°{self.temp_unit}/min\n' \
                            f'Interval: {interval3_temp:.2f}°{self.temp_unit} to {interval4_temp:.2f}°{self.temp_unit}'

                    plt.text(text_x, text_y, textstr, fontsize=10, color='blue',
                            verticalalignment='top', horizontalalignment='right',
                            bbox=dict(facecolor='white', alpha=0.5))
                    text_y -= 0.1 * (ylims[1] - ylims[0])  # Adjust vertical spacing

        plt.xlabel('Elapsed Time (s)')
        plt.ylabel(f'Temperature (°{self.temp_unit})')
        plt.title('Thermocouple Temperature Data')
        plt.tight_layout()  # Adjust layout to fit text
        plt.show()

    def export_to_excel(self):
        if not self.temp_data:
            QMessageBox.information(self, "No Data", "No temperature data to export.")
            return

        # Convert the temperature data to a DataFrame
        df = pd.DataFrame(self.temp_data)

        # Initialize cooling rate data
        cooling_rates_tc1 = {}
        cooling_rates_tc2 = {}

        # Check if cooling rate calculation is enabled and calculate accordingly
        if self.calculate_cooling_check_1.isChecked():
            interval1 = self.entry_interval_1.text()
            interval2 = self.entry_interval_2.text()
            cooling_rates_tc1 = self.calculate_cooling_rate_tc1(df, interval1, interval2)

        if self.calculate_cooling_check_2.isChecked():
            interval3 = self.entry_interval_3.text()
            interval4 = self.entry_interval_4.text()
            cooling_rates_tc2 = self.calculate_cooling_rate_tc2(df, interval3, interval4)

        # Pivot the DataFrame to get the desired format
        df_pivot = df.pivot_table(index='timestamp', columns='tc_id', values='temp_c').reset_index()
        df_pivot.columns.name = None  # Remove the index name for better readability
        df_pivot.rename(columns={'timestamp': 'Elapsed Time (s)'}, inplace=True)
        df_pivot.columns = [f'Thermocouple {int(col)} (°{self.temp_unit})' if isinstance(col, int) else col for col in df_pivot.columns]

        default_filename = f"TURTLE_Data_{datetime.now().strftime('%m-%d-%y_%H-%M-%S')}.xlsx"

        # Get the path to save the file with a default filename and location
        filename, _ = QFileDialog.getSaveFileName(self, "Save File", default_filename, "Excel Files (*.xlsx)")
        if filename:
            # Create an Excel writer object with xlsxwriter as the engine
            with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                df_pivot.to_excel(writer, sheet_name='Temperature Data', index=False)

                workbook = writer.book
                temperature_sheet = writer.sheets['Temperature Data']

                # Format columns for better visibility
                for col_num, value in enumerate(df_pivot.columns.values):
                    temperature_sheet.set_column(col_num, col_num, max(len(value), 15))

                # Create a line chart
                chart = workbook.add_chart({'type': 'line'})
                num_rows = len(df_pivot)

                # Add thermocouple data series to the chart
                tc_index = 1  # Thermocouple index counter
                for col_num in range(1, len(df_pivot.columns)):
                    col_letter = chr(65 + col_num)
                    chart.add_series({
                        'name': f'Thermocouple {tc_index}',
                        'categories': f'={temperature_sheet.name}!$A$2:$A${num_rows+1}',
                        'values': f'={temperature_sheet.name}!${col_letter}$2:${col_letter}${num_rows+1}',
                    })
                    tc_index += 1

                # Configure the chart layout
                chart.set_title({'name': 'Thermocouple Temperature Data'})
                chart.set_x_axis({'name': 'Elapsed Time (s)'})
                chart.set_y_axis({'name': f'Temperature (°{self.temp_unit})'})
                chart.set_legend({'position': 'bottom'})

                # Insert the chart to the right of the data (Column F)
                temperature_sheet.insert_chart(f'F2', chart)

                # Insert cooling rate information below the graph
                row_offset = num_rows + 5  # Ensure the information appears clearly below the graph
                temperature_sheet.write(row_offset, 5, 'Cooling Rate Information:')

                row = row_offset + 1
                if self.calculate_cooling_check_1.isChecked() and cooling_rates_tc1:
                    if 1 in cooling_rates_tc1 and cooling_rates_tc1[1]:  # Ensure data exists
                        data = cooling_rates_tc1[1]
                        temperature_sheet.write(row, 5, 'Thermocouple 1:')
                        temperature_sheet.write(row + 1, 5, f'Cooling Rate: {data["cooling_rate"]}°{self.temp_unit}/min')
                        temperature_sheet.write(row + 2, 5, f'Interval 1 Temp: {data["interval1_temp"]}°{self.temp_unit}')
                        temperature_sheet.write(row + 3, 5, f'Interval 2 Temp: {data["interval2_temp"]}°{self.temp_unit}')
                        temperature_sheet.write(row + 4, 5, f'Interval 1 Time: {data["interval1_time"]}s')
                        temperature_sheet.write(row + 5, 5, f'Interval 2 Time: {data["interval2_time"]}s')
                        row += 7  # Move to the next block for the next thermocouple

                if self.calculate_cooling_check_2.isChecked() and cooling_rates_tc2:
                    if 2 in cooling_rates_tc2 and cooling_rates_tc2[2]:  # Ensure data exists
                        data = cooling_rates_tc2[2]
                        temperature_sheet.write(row, 5, 'Thermocouple 2:')
                        temperature_sheet.write(row + 1, 5, f'Cooling Rate: {data["cooling_rate"]}°{self.temp_unit}/min')
                        temperature_sheet.write(row + 2, 5, f'Interval 1 Temp: {data["interval1_temp"]}°{self.temp_unit}')
                        temperature_sheet.write(row + 3, 5, f'Interval 2 Temp: {data["interval2_temp"]}°{self.temp_unit}')
                        temperature_sheet.write(row + 4, 5, f'Interval 1 Time: {data["interval1_time"]}s')
                        temperature_sheet.write(row + 5, 5, f'Interval 2 Time: {data["interval2_time"]}s')
                        row += 7  # Move to the next block

                if row == row_offset + 1:  # If no cooling rate data was written
                    temperature_sheet.write(row_offset + 2, 5, 'No cooling rate data available.')

            # Notify user of the successful export
            QMessageBox.information(self, "Export Success", f"Temperature data successfully exported to:\n{filename}")


    def find_arduino_port(self):
        ports = list_ports.comports()
        for port in ports:
            if "Arduino" in port.description:
                return port.device
            if "VID:PID=2341:0043" in port.hwid or "VID:PID=0403:6001" in port.hwid or "VID:PID=2341:0043" in port.hwid or "VID:PID=2A03:0010" in port.hwid or "VID:PID=1A86:7523" in port.hwid:
                return port.device
        return None

    def connect_to_arduino(self):
        if not self.connected:  # Check if already connected
            arduino_port = self.find_arduino_port()
            if arduino_port is not None:
                try:
                    self.ser = serial.Serial(arduino_port, 9600, timeout=1)
                    self.connected = True
                    self.send_to_arduino("TYPE:T;")
                    self.send_to_arduino("RATE:1;")

                except Exception as e:
                    QMessageBox.information(self, "Connection Error", f"Failed to connect to Arduino: {e}")
            else:
                QMessageBox.information(self, "Connection Error", "Arduino not found")

    def send_to_arduino(self, message):
        if self.connected and self.ser:
            try:
                self.ser.write(message.encode())
            except Exception as e:
                QMessageBox.information(self, "Error", f"Error sending data to Arduino. Error: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = ThermocoupleUI()
    window.show()
    sys.exit(app.exec())
