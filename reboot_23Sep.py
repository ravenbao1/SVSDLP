import time
import datetime
import tkinter as tk
import logging
from logging.handlers import RotatingFileHandler
import ctypes
import winreg
import wmi
import os
import argparse  # Import argparse for command-line arguments
from win11toast import toast
import subprocess

# Set up logging to a rotating file
log_file_path = r"C:\temp\weeklyreboot.log"
os.makedirs(os.path.dirname(log_file_path), exist_ok=True)

# Argument parsing to allow start and end time input along with day of the week
parser = argparse.ArgumentParser(description="Weekly Reboot Script")
parser.add_argument("--start", type=str, help="Countdown start time in 'Day HH:MM' format (24-hour clock, e.g., Wed 00:00)", default="Wed 00:00")
parser.add_argument("--end", type=str, help="Countdown end time in 'Day HH:MM' format (24-hour clock, e.g., Sat 10:00)", default="Sat 10:00")

args = parser.parse_args()

# Parse the start and end day/time from the arguments
try:
    start_day_str, start_time_str = args.start.split()
    end_day_str, end_time_str = args.end.split()

    # Convert the day names to weekday numbers (0=Monday, 6=Sunday)
    days_map = {'Mon': 0, 'Tue': 1, 'Wed': 2, 'Thu': 3, 'Fri': 4, 'Sat': 5, 'Sun': 6}
    START_DAY = days_map[start_day_str]
    END_DAY = days_map[end_day_str]

    # Convert the time strings to time objects
    COUNTDOWN_START_TIME = datetime.datetime.strptime(start_time_str, "%H:%M").time()
    COUNTDOWN_END_TIME = datetime.datetime.strptime(end_time_str, "%H:%M").time()

except ValueError:
    print("Invalid format provided for --start or --end. Use 'Day HH:MM' format, e.g., 'Wed 00:00'.")
    exit(1)

# Configurable additional checkpoints (add tuples of (hour, minute) as needed)
ADDITIONAL_CHECKPOINTS = [(COUNTDOWN_START_TIME.hour, COUNTDOWN_START_TIME.minute),
                          (COUNTDOWN_END_TIME.hour, COUNTDOWN_END_TIME.minute)]

# Configuring the logger
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

# Create a rotating file handler that backs up the file once it reaches 5MB, keeping 5 backup files.
handler = RotatingFileHandler(log_file_path, maxBytes=5 * 1024 * 1024, backupCount=5)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)

# Initialize Windows 11 Toast Notification
def show_notification(time_left):
    days, remainder = divmod(time_left.total_seconds(), 86400)
    hours, remainder = divmod(remainder, 3600)
    minutes, seconds = divmod(remainder, 60)
    message = f"Device will reboot in {int(days):02d} days, {int(hours):02d} hours, {int(minutes):02d} minutes, and {int(seconds):02d} seconds."

    logging.debug(f"Showing notification: {message}")
    toast("Weekly Reboot Program", message, duration="short", app_id="OCBC")

# Registry path and key
REG_PATH = r"SOFTWARE\WeeklyReboot"
REG_NAME = "rebooted"

def read_or_create_registry_value(current_uptime, in_countdown, countdown_start):
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, REG_PATH, 0, winreg.KEY_READ) as key:
            value, _ = winreg.QueryValueEx(key, REG_NAME)
            return value
    except FileNotFoundError:
        if in_countdown and current_uptime is not None:
            last_reboot_time = datetime.datetime.now() - datetime.timedelta(seconds=current_uptime)
            if last_reboot_time >= countdown_start:
                logging.debug(f"Registry key not found. Creating {REG_NAME} and setting it to 1 due to reboot during countdown period.")
                write_registry_value(1)
                return 1
        logging.debug(f"Registry key not found. Creating {REG_NAME} and setting it to 0.")
        write_registry_value(0)
        return 0

def write_registry_value(value):
    try:
        with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, REG_PATH) as key:
            winreg.SetValueEx(key, REG_NAME, 0, winreg.REG_DWORD, value)
    except Exception as e:
        logging.error(f"Failed to write to registry: {e}")

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

def format_uptime(uptime_seconds):
    days, remainder = divmod(uptime_seconds, 86400)
    hours, remainder = divmod(remainder, 3600)
    minutes, _ = divmod(remainder, 60)
    return f"{int(days)}d {int(hours)}h {int(minutes)}m"

def time_until_reboot(now):
    if now.weekday() == END_DAY:
        reboot_datetime = datetime.datetime.combine(now.date(), COUNTDOWN_END_TIME)
        return reboot_datetime - now

    reboot_datetime = datetime.datetime.combine(now.date(), COUNTDOWN_END_TIME)
    while reboot_datetime.weekday() != END_DAY:
        reboot_datetime += datetime.timedelta(days=1)

    return reboot_datetime - now

def get_last_and_next_countdown_periods(now):
    last_countdown_end = now.replace(hour=COUNTDOWN_END_TIME.hour, 
                                     minute=COUNTDOWN_END_TIME.minute, 
                                     second=COUNTDOWN_END_TIME.second, 
                                     microsecond=0)
    
    while last_countdown_end.weekday() != END_DAY or last_countdown_end >= now:
        last_countdown_end -= datetime.timedelta(days=1)
        
    last_countdown_start = last_countdown_end - datetime.timedelta(days=((END_DAY - START_DAY) % 7))
    last_countdown_start = last_countdown_start.replace(hour=COUNTDOWN_START_TIME.hour, 
                                                       minute=COUNTDOWN_START_TIME.minute, 
                                                       second=COUNTDOWN_START_TIME.second)

    next_countdown_start = now.replace(hour=COUNTDOWN_START_TIME.hour, 
                                       minute=COUNTDOWN_START_TIME.minute, 
                                       second=COUNTDOWN_START_TIME.second, 
                                       microsecond=0)

    while next_countdown_start.weekday() != START_DAY or next_countdown_start <= now:
        next_countdown_start += datetime.timedelta(days=1)

    next_countdown_end = next_countdown_start + datetime.timedelta(days=((END_DAY - START_DAY) % 7))
    next_countdown_end = next_countdown_end.replace(hour=COUNTDOWN_END_TIME.hour, 
                                                   minute=COUNTDOWN_END_TIME.minute, 
                                                   second=COUNTDOWN_END_TIME.second)
    
    return last_countdown_start, last_countdown_end, next_countdown_start, next_countdown_end

def is_in_countdown_period(now):
    start_date = now - datetime.timedelta(days=(now.weekday() - START_DAY) % 7)
    start_datetime = datetime.datetime.combine(start_date, COUNTDOWN_START_TIME)

    end_date = now - datetime.timedelta(days=(now.weekday() - END_DAY) % 7)
    end_datetime = datetime.datetime.combine(end_date, COUNTDOWN_END_TIME)

    if end_datetime <= start_datetime:
        end_datetime += datetime.timedelta(weeks=1)

    return start_datetime <= now <= end_datetime, start_datetime

# Small transparent countdown gadget with no background
def run_countdown_gadget(time_left):
    root = tk.Tk()
    root.title("Reboot Countdown")

    user32 = ctypes.windll.user32
    screen_width = user32.GetSystemMetrics(0)
    screen_height = user32.GetSystemMetrics(1)

    window_width = 500  
    window_height = 100
    x_position = screen_width - window_width - 20  
    y_position = 10  

    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
    root.overrideredirect(True)  

    root.wm_attributes("-transparentcolor", "black")  

    light_red = "#ff6666"  

    label = tk.Label(root, text="", font=("Helvetica", 24, "bold"), fg=light_red, bg="black")
    label.pack(expand=True, fill='both')

    def update_label():
        nonlocal time_left
        days, remainder = divmod(time_left.total_seconds(), 86400)
        hours, remainder = divmod(remainder, 3600)
        minutes, seconds = divmod(remainder, 60)

        label.config(text=f"Reboot in: {int(days):02d}d {int(hours):02d}h {int(minutes):02d}m {int(seconds):02d}s")
        
        if time_left.total_seconds() <= 0:
            root.destroy()
        else:
            time_left -= datetime.timedelta(seconds=1)
            label.after(1000, update_label)

    update_label()
    root.lower()
    root.mainloop()

# Main function
def main():
    last_uptime, last_boot_time = get_system_uptime()
    perform_check_on_startup = True

    now = datetime.datetime.now()
    last_start, last_end, next_start, next_end = get_last_and_next_countdown_periods(now)
    
    logging.info(f"Last countdown period: Start - {last_start}, End - {last_end}")
    logging.info(f"Next countdown period: Start - {next_start}, End - {next_end}")

    while True:
        now = datetime.datetime.now()
        
        in_countdown, countdown_start = is_in_countdown_period(now)
        countdown_end = countdown_start.replace(hour=COUNTDOWN_END_TIME.hour, 
                                                minute=COUNTDOWN_END_TIME.minute, 
                                                second=COUNTDOWN_END_TIME.second)

        current_uptime, current_boot_time = get_system_uptime()
        
        if current_uptime is not None:
            logging.debug(f"System uptime: {format_uptime(current_uptime)}")

        if in_countdown:
            reboot_flag = read_or_create_registry_value(current_uptime, in_countdown, countdown_start)
            if reboot_flag == 1:
                logging.debug("Skipping countdown as reboot already occurred during countdown period.")
            else:
                time_left = time_until_reboot(now)
                show_notification(time_left)
                run_countdown_gadget(time_left)

        if perform_check_on_startup or \
           (now.weekday() == END_DAY and (now.hour, now.minute, now.second) == (COUNTDOWN_END_TIME.hour, COUNTDOWN_END_TIME.minute, COUNTDOWN_END_TIME.second)) or \
           (now.weekday() == START_DAY and (now.hour, now.minute) in ADDITIONAL_CHECKPOINTS):
           
            logging.debug("Performing additional check.")
            perform_check_on_startup = False

            if not in_countdown:
                last_reboot_time = now - datetime.timedelta(seconds=current_uptime)
                logging.debug(f"Last reboot time: {last_reboot_time}")

                if countdown_start <= last_reboot_time <= countdown_end:
                    logging.debug("The device was last rebooted during the previous countdown period. No reboot required.")
                    write_registry_value(1)
                elif last_reboot_time > countdown_end:
                    logging.debug("The device has already been rebooted after the last countdown period. No additional reboot required.")
                    write_registry_value(1)
                else:
                    if now < countdown_end:
                        logging.debug("It's still before the end of the current countdown period. No reboot required yet.")
                    else:
                        subprocess.run(["cmd", "/c", "shutdown /r /f /t 0"], check=True)
                        logging.info("The device is bring rebooted.")
                        break

        # Adjust sleep time
        if in_countdown or (countdown_start - datetime.timedelta(minutes=1) <= now <= countdown_end + datetime.timedelta(minutes=1)):
            time.sleep(1)  # Check every second within 1 minute before and after countdown start and end
        else:
            time.sleep(60)  # Check every 60 seconds outside these times

if __name__ == "__main__":
    main()
