import time
import datetime
import tkinter as tk
import logging
from logging.handlers import RotatingFileHandler
import ctypes
import winreg
import wmi
import os
import threading
import gc

# Set up logging to a rotating file
log_file_path = r"C:\temp\weeklyreboot.log"
os.makedirs(os.path.dirname(log_file_path), exist_ok=True)

def calculate_difference(startDay, startHour, endDay, endHour):
    # Calculate the day difference
    if endDay >= startDay:
        dayDifference = endDay - startDay
    else:
        dayDifference = (endDay + 7) - startDay

    # Calculate the hour difference
    hourDifference = endHour - startHour

    # Adjust if the hour difference is negative
    if hourDifference < 0:
        dayDifference -= 1
        hourDifference += 24

    return dayDifference, hourDifference

startDay = 2
endDay = 5
startHour = 0
endHour = 22
diffDay, diffHour = calculate_difference(startDay,startHour,endDay, endHour)

# Configuring the logger
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

# Global variables to manage the countdown gadget
gadget_running = False
gadget_lock = threading.Lock()

# Create a rotating file handler that backs up the file once it reaches 5MB, keeping 5 backup files.
handler = RotatingFileHandler(log_file_path, maxBytes=5 * 1024 * 1024, backupCount=5)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)

# Registry path and key
REG_PATH = r"SOFTWARE\WeeklyReboot"
REG_NAME = "rebooted"

# Function to read or create DWORD value from the registry
def read_or_create_registry_value(current_uptime, in_countdown, countdown_start):
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, REG_PATH, 0, winreg.KEY_READ) as key:
            value, _ = winreg.QueryValueEx(key, REG_NAME)
            logging.debug(f"Read registry value: {value}")
            return value
    except FileNotFoundError:
        # If the registry key doesn't exist, create it
        if in_countdown and current_uptime is not None:
            last_reboot_time = datetime.datetime.now() - datetime.timedelta(seconds=current_uptime)
            if last_reboot_time >= countdown_start:
                logging.debug(f"Registry key not found. Creating {REG_NAME} and setting it to 1 due to reboot during countdown period.")
                write_registry_value(1)
                return 1
        logging.debug(f"Registry key not found. Creating {REG_NAME} and setting it to 0.")
        write_registry_value(0)
        return 0


# Function to write DWORD value to the registry
def write_registry_value(value):
    try:
        # Open the key with write access if it exists, otherwise create it
        with winreg.CreateKeyEx(winreg.HKEY_LOCAL_MACHINE, REG_PATH, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, REG_NAME, 0, winreg.REG_DWORD, value)
        logging.debug(f"Successfully set {REG_NAME} to {value} in the registry.")
    except Exception as e:
        logging.error(f"Failed to write to registry: {e}")

# Function to get system uptime using WMI
def get_system_uptime():
    try:
        c = wmi.WMI()
        for os in c.Win32_OperatingSystem():
            last_boot_time = os.LastBootUpTime.split('.')[0]
            last_boot_datetime = datetime.datetime.strptime(last_boot_time, '%Y%m%d%H%M%S')
            uptime_seconds = (datetime.datetime.now() - last_boot_datetime).total_seconds()
            return uptime_seconds, last_boot_datetime
    except Exception as e:
        logging.error(f"Error detecting uptime: {e}")
        return None, None

# Function to format uptime into days, hours, and minutes
def format_uptime(uptime_seconds):
    days, remainder = divmod(uptime_seconds, 86400)
    hours, remainder = divmod(remainder, 3600)
    minutes, _ = divmod(remainder, 60)
    return f"{int(days)}d {int(hours)}h {int(minutes)}m"

# Function to calculate time remaining until reboot
def time_until_reboot(now):
    reboot_day = endDay  # Saturday
    reboot_time = datetime.time(endHour, 00)  # Updated to match end of countdown

    reboot_datetime = datetime.datetime.combine(now.date(), reboot_time)
    while reboot_datetime.weekday() != reboot_day:
        reboot_datetime += datetime.timedelta(days=1)

    return reboot_datetime - now

# Function to check if we are in the countdown period (Wed 00:00 to Sat 10:00)
def is_in_countdown_period(now):
    countdown_start = now.replace(hour=startHour, minute=0, second=0, microsecond=0)
    # To be updated
    while countdown_start.weekday() != startDay:  # Monday
        countdown_start -= datetime.timedelta(days=1)
    # To be updated
    countdown_end = countdown_start + datetime.timedelta(days=diffDay, hours=diffHour)  # Monday 16:00

    return countdown_start <= now <= countdown_end, countdown_start, countdown_end

def get_countdown_periods(now):
    """
    Determine the last and next countdown periods.
    - Countdown starts on Wednesday at 10:00
    - Countdown ends on Saturday at 16:00
    """
    # Calculate the last countdown start (Wednesday 10:00)
    last_countdown_start = now.replace(hour=startHour, minute=0, second=0, microsecond=0)
    if (last_countdown_start > now) and (last_countdown_start.weekday() == startDay):
        last_countdown_start = last_countdown_start - datetime.timedelta(weeks=1)
    elif last_countdown_start.weekday() != startDay:
        while last_countdown_start.weekday() != startDay:  # 5 = Saturday
            last_countdown_start -= datetime.timedelta(days=1)

    # Calculate the last countdown end (Saturday 16:00)
    last_countdown_end = last_countdown_start + datetime.timedelta(days=(endDay - startDay), hours=(endHour - startHour))
    
    # Calculate the next countdown period by adding one week
    next_countdown_start = last_countdown_start + datetime.timedelta(weeks=1)
    next_countdown_end = last_countdown_end + datetime.timedelta(weeks=1)

    return last_countdown_start, last_countdown_end, next_countdown_start, next_countdown_end

# Small transparent countdown gadget with no background
def run_countdown_gadget(time_left):
    global gadget_running

    with gadget_lock:
        if gadget_running:
            logging.debug("Countdown gadget is already running. Skipping creation.")
            return
        gadget_running = True

    root = tk.Tk()
    root.title("Reboot Countdown")
    try:
        # Get screen width and height to position the window in the top-right corner
        user32 = ctypes.windll.user32
        screen_width = user32.GetSystemMetrics(0)
        screen_height = user32.GetSystemMetrics(1)

        # Set the window size and move it to the top-right corner
        window_width = 500  # Increased width to fit the full text
        window_height = 100

        # Calculate the x_position and y_position to center the window on the screen
        x_position = (screen_width // 2) - (window_width // 2)
        y_position = -10

        root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        root.overrideredirect(True)  # Remove window decorations

        # Set window transparency (1.0 is fully opaque, 0.0 is fully transparent)
        root.wm_attributes("-transparentcolor", "black")  # Make black transparent

        # Lighter red color towards pink for the text with an effect of 30% opacity
        light_red = "#ff6666"  # This is a lighter red that approximates 30% opacity

        # Label to display the countdown message with no background
        label = tk.Label(root, text="", font=("Helvetica", 24, "bold"), fg=light_red, bg="black")
        label.pack(expand=True, fill='both')

        # Function to update the label text with the "Reboot in: xxx" format, ensuring 2 digits for day, hour, min, sec
        def update_label():
            # Dynamically recalculate the time left
            now = datetime.datetime.now()
            reboot_datetime = datetime.datetime.combine(now.date(), datetime.time(endHour, 0))
            
            # Ensure the reboot date aligns with the correct weekday
            while reboot_datetime.weekday() != endDay:
                reboot_datetime += datetime.timedelta(days=1)

            # Calculate time remaining
            time_left = reboot_datetime - now

            # Ensure time doesn't go negative
            if time_left.total_seconds() <= 0:
                label.config(text="Rebooting now...")
                root.destroy()
                return

            # Update the label text with formatted time
            days, remainder = divmod(time_left.total_seconds(), 86400)
            hours, remainder = divmod(remainder, 3600)
            minutes, seconds = divmod(remainder, 60)
            label.config(text=f"Reboot in: {int(days):02d}d {int(hours):02d}h {int(minutes):02d}m {int(seconds):02d}s")

            # Schedule the next update
            label.after(1000, update_label)

        update_label()

        def on_close():
            global gadget_running
            gadget_running = False
            root.destroy()

        # Safely set the protocol handler
        #if root.winfo_exists():
        #    root.protocol("WM_DELETE_WINDOW", on_close)

        root.mainloop()

    except Exception as e:
        logging.error(f"Error in countdown gadget: {e}")
    finally:
        with gadget_lock:
            gadget_running = False


# Main function
def main():
    last_uptime, last_boot_time = get_system_uptime()  # Store the last system uptime and boot time

    while True:
        now = datetime.datetime.now()

        # Check if we are in the countdown period
        in_countdown, countdown_start, countdown_end = is_in_countdown_period(now)
        current_uptime, current_boot_time = get_system_uptime()

        if current_uptime is not None:
            logging.debug("xxxxxxxxxxxxxxxxxxx New Record xxxxxxxxxxxxxxxxxxx")
            logging.debug(f"System uptime: {format_uptime(current_uptime)}")
            # Get the last reboot time
            last_reboot_time = datetime.datetime.now() - datetime.timedelta(seconds=current_uptime)
            logging.debug(f"Last reboot time: {last_reboot_time}")

        if in_countdown:
            # If the last reboot happened during the countdown period, set the flag to 1
            last_reboot_time = datetime.datetime.now() - datetime.timedelta(seconds=current_uptime)
            if last_reboot_time >= countdown_start:
                logging.debug(f"Currently in countdown period from {countdown_start} to {countdown_end}.")
                logging.debug("Last reboot during countdown period. Setting reboot flag to 1.")
                write_registry_value(1)
            else:
                logging.debug(f"Currently in countdown period from {countdown_start} to {countdown_end}.")
                logging.debug("Last reboot outside countdown period. Setting reboot flag to 0.")
                write_registry_value(0)
            
            # Read or create the reboot flag from the registry
            reboot_flag = read_or_create_registry_value(current_uptime, in_countdown, countdown_start)

            # If reboot flag is 1, skip the countdown
            if reboot_flag == 1:
                logging.debug("Skipping countdown as reboot already occurred during countdown period.")
            else:
                # If reboot flag is 0, start countdown
                time_left = time_until_reboot(now)
                logging.debug(f"Rebooting in {time_left}.")
                # Start the countdown gadget in a new thread only if it's not already running
                if not gadget_running:
                    threading.Thread(target=run_countdown_gadget, args=(time_left,), daemon=True).start()
                    logging.debug("gadget started")
                write_registry_value(0)
        else:
            # Outside countdown period
            logging.debug("Outside countdown period.")
            
            # Get the last and next countdown periods
            last_countdown_start, last_countdown_end, next_countdown_start, next_countdown_end = get_countdown_periods(now)

            logging.debug(f"Last countdown period: {last_countdown_start} to {last_countdown_end}")
            logging.debug(f"Next countdown period: {next_countdown_start} to {next_countdown_end}")            

            if last_reboot_time < last_countdown_start:
                # No reboot detected during the last countdown period, force reboot
                logging.debug("No reboot detected during the last countdown period. Rebooting immediately.")
                write_registry_value(0)
                os.system("shutdown /r /t 5")
            else:
                # Reset reboot flag to 0 if outside the countdown period
                logging.debug("Reboot already performed during or after the last countdown period.")
                write_registry_value(0)

        # Sleep for 1 min before checking again
        time.sleep(60)
        gc.collect()

if __name__ == "__main__":
    main()
