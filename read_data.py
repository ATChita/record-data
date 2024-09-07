import subprocess
import pandas as pd
import time
import datetime

# Function to read data from the log file
def read_data_from_log(log_file_path):
    data = []
    try:
        with open(log_file_path, "r") as file:
            for line in file:
                if "X :" in line:
                    parts = line.strip().split(',')
                    try:
                        timestamp = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3] 
                        x = float(parts[0].split(':')[1].strip())
                        y = float(parts[1].split(':')[1].strip())
                        z = float(parts[2].split(':')[1].strip())
                        vector = float(parts[3].split(':')[1].strip())
                        data.append([timestamp, x, y, z, vector])
                    except IndexError:
                        print("Skipping line due to format issue:", line)
    except FileNotFoundError:
        print(f"File not found: {log_file_path}")
    return data

log_file_path = "C:/Users/ASUS TUF/Documents/py/Jlink_view/-LogToFile"
output_excel_path = "C:/Users/ASUS TUF/Documents/py/Jlink_view/kdata.xlsx"
duration = 1200  # Duration in second

# Start the JLink RTT logger
cmd = [
        "JLinkRTTLogger", 00
        "-Device", "nRF52840_xxAA", 
        "-IF", "SWD", 
        "-Speed", "4000", 
        "-RTTChannel", "0", 
        "-LogToFile", log_file_path
    ]

try:
    process = subprocess.Popen(cmd)
    print("JLink RTT Logger started. waiting for duration")

    # wait for log data in logfile 
    time.sleep(duration)

    # Terminate the process after the wait
    process.terminate()
    print("JLink RTT Logger stopped")

    # Process the read data from log file
    data = read_data_from_log(log_file_path)

    if data:
        df = pd.DataFrame(data, columns=["time", "X", "Y", "Z", "Vector"])
                
        # Save the data to Excel
        df.to_excel(output_excel_path, index=False, engine='openpyxl')
        print("Data has been processed and saved to kdata.xlsx")
                
    else:
        print("No data to save")

except Exception as e:
    print(f"An error occurred: {e}")

# Optionally wait before starting the next iteration
time.sleep(1)  # Adjust the sleep time as needed