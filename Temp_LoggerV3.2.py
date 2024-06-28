from serial import Serial
from serial.tools import list_ports
import pandas as pd
from matplotlib.pyplot import figure, plot, xlabel, ylabel, title, legend, show
from datetime import datetime
from threading import Thread
from time import sleep
from tkinter import messagebox, DISABLED, NORMAL
from customtkinter import CTk, CTkFrame, CTkLabel, CTkButton, CTkOptionMenu, set_appearance_mode
import numpy as np
import os
from subprocess import call
from sys import exit
import xlsxwriter

class ThermocoupleUI(CTk):
    def __init__(self):
        super().__init__()

        self.connected = False
        self.ser = None

        icon_filename = "AGGRC_logo.ico"
        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, icon_filename)
        self.iconbitmap(icon_path)

        # default variables
        self.start_time = None
        self.is_recording = False
        self.record_thread = None
        self.temp_data = []
        self.reading_thread = None  # Thread for reading temperature data continuously
        self.first_connection = True
        self.timestamp_counter = 0
        self.sampling_rate = 1

        # Thermocouple temperature labels
        self.temps_f = []
        self.temps_c = []

        # connection status
        self.connection_status_labels = []

        # auto connects to Arduino COM port
        self.connect_to_arduino()

        # Layout setup
        self.setup_ui()

        # disables start button until at least 1 thermocouple detected
        self.start_button.configure(state=DISABLED)

        # start reading temperature data
        self.start_reading_data()

        # closing protocol (closes serial port before closing application)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def setup_ui(self):
        self.title("AGGRC Temperature Logger")
        self.geometry("1000x460")  # Adjusted window size for a more compact UI

        # Set the theme to dark
        set_appearance_mode("dark")

        # Main frame for thermocouples and their controls
        main_frame = CTkFrame(self)
        main_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        # Thermocouple display frame, using grid layout for precise placement
        tc_frame = CTkFrame(main_frame)
        tc_frame.place(relx=0, rely=0, relwidth=1)

        # Define a larger font size for better visibility
        large_font = ("Arial", 16)

        # tc displays
        for i in range(4):
            row = (i // 2)
            column = i % 2
            temp_frame = CTkFrame(tc_frame, corner_radius=10)
            temp_frame.grid(row=row, column=column, padx=10, pady=10, sticky="nsew")

            # Ensure temp_frame takes up available space (Configure grid column weight)
            tc_frame.grid_columnconfigure(column, weight=1)  # Adjust weight to allow expansion

            # Connection status label with rounded corners, specified font and colors
            connection_status_label = CTkLabel(temp_frame, text="Not Connected", font=large_font, fg_color="#808080", bg_color="#333", corner_radius=10)
            connection_status_label.pack(pady=(5, 0))
            self.connection_status_labels.append(connection_status_label)

            # Thermocouple label setup, now with padx to take up more horizontal space
            CTkLabel(temp_frame, text=f"Thermocouple {i+1}", anchor="center", font=large_font).pack(pady=2)
            self.temps_f.append(CTkLabel(temp_frame, text="N/A °F", font=large_font))
            self.temps_c.append(CTkLabel(temp_frame, text="N/A °C", font=large_font))

            # Adjust padding to ensure labels take up more space and are centered
            self.temps_f[i].pack(pady=10, padx=20)
            self.temps_c[i].pack(pady=10, padx=20)

        # Frame for control buttons and elapsed time, positioned to the right
        control_frame = CTkFrame(self, width=180)
        control_frame.pack(side="right", fill="y", padx=15, pady=10)

        # start button and elapsed time counter
        self.start_button = CTkButton(control_frame, text="Start Recording", command=self.toggle_recording, fg_color="green", font=large_font)
        self.start_button.pack(pady=10)
        self.elapsed_time_var = CTkLabel(control_frame, text="Time Elapsed: 0s", font=large_font)
        self.elapsed_time_var.pack()

        # Options frame for sampling rate and type
        options_frame = CTkFrame(main_frame)
        options_frame.pack(side='bottom', pady=(50, 0), fill='x')

        sampling_rate_options = [f"Sample every {i} second(s)" for i in range(1, 6)]
        self.sampling_rate_var = CTkOptionMenu(options_frame, values=sampling_rate_options, command=self.update_sampling_rate, font=large_font)
        self.sampling_rate_var.set("Sample every 1 second(s)")
        self.sampling_rate_var.pack(padx=10, pady=10)

        thermocouple_type_options = ["Type " + type for type in ["K", "J", "T", "E", "N", "S", "R", "B"]]
        self.thermocouple_type_var = CTkOptionMenu(options_frame, values=thermocouple_type_options, font=large_font, command=self.update_tc_type)
        self.thermocouple_type_var.set("Type K")
        self.thermocouple_type_var.pack(padx=10, pady=10)

        # Additional buttons for Graph and Export to Excel
        CTkButton(control_frame, text="Graph Last Data Points", command=self.show_graph, font=large_font).pack(padx=2, pady=10)  # Adjusted padding
        CTkButton(control_frame, text="Export Last Data to Excel", command=self.export_to_excel, font=large_font).pack(padx=0, pady=10)  # Adjusted padding

    def update_sampling_rate(self, rate):
        # Assume rate comes in as "Sample every X second(s)"
        rate_value = rate.split(" ")[2]  # Extract the X
        self.send_to_arduino(f"RATE:{rate_value};")
        self.sampling_rate = int(rate_value)

    def update_tc_type(self, type):
        # comes in is Type X
        tc_type = type.split(" ")[1]  # extracts X
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
                            temp_data = statuses.split(",")[:-1]  # Exclude the last empty part after split

                            data_collected = []  # Temporary list to store collected data

                            for data in temp_data:
                                tc, status = data.split(":")
                                index = int(tc[1]) - 1  # Convert 'T1' to 0, 'T2' to 1, etc.
                                if status == "Not Connected":
                                    self.connection_status_labels[index].configure(text="Not Connected", fg_color="#808080")
                                    self.temps_f[index].configure(text="N/A °F")
                                    self.temps_c[index].configure(text="N/A °C")
                                else:
                                    if self.first_connection:
                                        self.start_button.configure(state=NORMAL)
                                        self.first_connection = False
                                    self.connection_status_labels[index].configure(text="Connected", fg_color="green")
                                    temp_c = float(status)
                                    temp_f = (1.8 * temp_c) + 32
                                    self.temps_f[index].configure(text=f"{temp_f:.2f}°F")
                                    self.temps_c[index].configure(text=f"{temp_c:.2f}°C")

                                    # Collect data for later timestamp increment
                                    if self.is_recording:
                                        data_collected.append({
                                            'tc_id': index + 1,  # Thermocouple ID (assuming it starts from 1)
                                            'temp_c': temp_c
                                        })
                            # Update timestamp and add collected data to temp_data
                            if self.is_recording and data_collected:
                                self.update_elapsed_time()
                                for reading in data_collected:
                                    self.current_reading = {
                                        'timestamp': self.timestamp_counter,  # Use timestamp counter
                                        'tc_id': reading['tc_id'],
                                        'temp_c': reading['temp_c']
                                    }
                                    self.temp_data.append(self.current_reading)
                                self.timestamp_counter += self.sampling_rate  # Increment by sampling rate only once per complete read cycle
                except Exception as e:
                    messagebox.showinfo("Error", f"Error Reading from Arduino. Error: {e}")

            sleep(0.2)  # Adjust sleep time as needed

    def toggle_recording(self):
        if self.is_recording:
            # Stop recording
            self.is_recording = False
            self.start_button.configure(text="Start Recording", fg_color="green")
            self.timestamp_counter = 0
        else:
            # Clear temperature data and reset start time for the new experiment
            self.temp_data.clear()
            self.start_time = datetime.now()

            # Start recording
            self.is_recording = True
            self.start_button.configure(text="Stop Recording", fg_color="red")
            self.update_elapsed_time()

    def update_elapsed_time(self):
        if self.is_recording and self.start_time:
            current_time_elapsed = (datetime.now() - self.start_time).total_seconds()
            elapsed_time = current_time_elapsed
            self.elapsed_time_var.configure(text=f"Time Elapsed: {int(elapsed_time)}s")

    def show_graph(self):
        if not self.temp_data:
            messagebox.showinfo("No Data", "No temperature data to plot.")
            return

        # Convert to DataFrame
        df = pd.DataFrame(self.temp_data)

        figure(figsize=(10, 6))

        for tc_id in df['tc_id'].unique():
            tc_df = df[df['tc_id'] == tc_id]
            plot(tc_df['timestamp'], tc_df['temp_c'], label=f'Thermocouple {tc_id}')

        xlabel('Time (s)')
        ylabel('Temperature (°C)')
        title('Temperature Readings Over Time')
        legend()
        show()

    def export_to_excel(self):
        if not self.temp_data:
            messagebox.showinfo("No Data", "No temperature data to export.")
            return

        df = pd.DataFrame(self.temp_data)

        # Pivot the DataFrame to get the desired format
        df_pivot = df.pivot(index='timestamp', columns='tc_id', values='temp_c').reset_index()
        df_pivot.columns.name = None  # Remove the index name for better readability
        df_pivot.rename(columns={'timestamp': 'Time (s)'}, inplace=True)
        df_pivot.columns = [f'Thermocouple {int(col)} (°C)' if isinstance(col, int) else col for col in df_pivot.columns]

        # Get the path to the desktop
        desktop_path = path.join(path.expanduser("~"), "Desktop")

        # Create the "TURTLE Data" folder on the desktop if it doesn't exist
        turtle_data_folder = path.join(desktop_path, "TURTLE Data")
        makedirs(turtle_data_folder, exist_ok=True)

        file_name = f"Temperature_Data_{datetime.now().strftime('%m-%d-%Y_%H-%M-%S')}.xlsx"
        file_path = path.join(turtle_data_folder, file_name)

        # Export the pivoted DataFrame to Excel
        df_pivot.to_excel(file_path, index=False, engine='xlsxwriter')

        # Display the full file path in a message box
        messagebox.showinfo("Export Successful", f"Temperature data has been exported to:\n{file_path}")

    def find_arduino_port(self):
        ports = list_ports.comports()
        for port in ports:
            if "Arduino" in port.description:
                return port.device
            # VID:PID for Arduino Nano and clones
            elif "VID:PID=2341:0043" in port.hwid or "VID:PID=0403:6001" in port.hwid or "VID:PID=2341:0043" in port.hwid or "VID:PID=2A03:0010" in port.hwid:
                return port.device
        return None

    def connect_to_arduino(self):
        if not self.connected:  # Check if already connected
            arduino_port = self.find_arduino_port()
            if arduino_port is not None:
                try:
                    self.ser = Serial(arduino_port, 9600, timeout=1)
                    self.connected = True
                    # messagebox.showinfo("Connected",f"Arduino connected on port: {arduino_port}")
                    # Set initial thermocouple type ('K') and sampling rate ('1')
                    self.send_to_arduino("TYPE:K;")
                    self.send_to_arduino("RATE:1;")
                except Exception as e:
                    messagebox.showinfo("Connection Error", f"Failed to connect to Arduino: {e}")
            else:
                messagebox.showinfo("Connection Error", "Arduino not found")

    def send_to_arduino(self, message):
        if self.connected and self.ser:
            self.ser.write(message.encode())

    def on_close(self):
        if self.ser:
            self.ser.close()
        self.destroy()

if __name__ == "__main__":
    app = ThermocoupleUI()
    app.mainloop()
