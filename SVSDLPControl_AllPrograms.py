import logging
from logging.handlers import RotatingFileHandler
import time
import win32gui
import win32process
import psutil
import winreg
import keyboard
import gc
import os

# Setup logging
log_file_path = 'C:\\Windows\\SVSDLPControl.log'
log_dir = os.path.dirname(log_file_path)

# Create the directory if it doesn't exist
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

# Configure the rotating file handler
handler = RotatingFileHandler(log_file_path, maxBytes=5*1024*1024, backupCount=5)  # 5MB limit and keep 5 backup files
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s', handlers=[handler])
logger = logging.getLogger(__name__)

# Registry paths and values
explorer_key_path = r"Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
disallow_run_key = "DisallowRun"
snipping_tool_value_name = "SnippingTool"
snipping_tool_value_data = "SnippingTool.exe"
steps_recorder_value_name = "StepsRecorder"
steps_recorder_value_data = "psr.exe"

# Keywords to check for
browser_keywords = {
    "msedge.exe": [
        "SignPlus for OCBC Bank",
        "ocbc retrieval",
        "ocbc po",
        "SignPlus for OCBC Malaysia",
    ],
    "chrome.exe": [
        "SignPlus for OCBC Bank",
        "ocbc retrieval",
        "ocbc po",
        "SignPlus for OCBC Malaysia"
    ],
    "iexplore.exe": [
        "SignPlus for OCBC Bank",
        "ocbc retrieval",
        "ocbc po",
        "SignPlus for OCBC Malaysia",
    ]
}
other_keywords = [
    "edit account",
    "new account",
    "signplus41_signcheck",
    "new signatory",
    "edit signatory",
    "reject the cheque",
    "rules -",
    "search"
]

# Flag to indicate if the Print Screen key should be blocked
block_print_screen = False

# Define the hook procedure
def on_print_screen(event):
    global block_print_screen
    if block_print_screen and event.event_type == 'down':
        return False  # Block the key press

def install_keyboard_hooks():
    try:
        keyboard.hook_key('print screen', on_print_screen, suppress=True)
        logger.info("Installed Print Screen key hook")
    except Exception as e:
        logger.error(f"Error installing keyboard hooks: {e}")

def uninstall_keyboard_hooks():
    try:
        keyboard.unhook_all()
        logger.info("Uninstalled all keyboard hooks")
    except Exception as e:
        logger.error(f"Error uninstalling keyboard hooks: {e}")

def get_all_window_titles():
    titles = []

    def enum_window_callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title:
                titles.append((hwnd, title.strip()))

    win32gui.EnumWindows(enum_window_callback, None)
    return titles

def get_process_info(hwnd):
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    try:
        process = psutil.Process(pid)
        return process.name().lower(), pid  # Make process name lowercase
    except psutil.NoSuchProcess:
        return None, None

def is_window_minimized(hwnd):
    return win32gui.IsIconic(hwnd)

def set_registry_value_for_all_users(key, value_name, value_data):
    user_hive = winreg.HKEY_USERS
    user_profiles = []

    try:
        i = 0
        while True:
            user_profiles.append(winreg.EnumKey(user_hive, i))
            i += 1
    except OSError:
        pass  # No more subkeys

    for profile in user_profiles:
        if profile not in (".DEFAULT", "S-1-5-18", "S-1-5-19", "S-1-5-20"):
            try:
                user_key_path = f"{profile}\\{key}"
                with winreg.CreateKeyEx(user_hive, user_key_path, 0, winreg.KEY_ALL_ACCESS) as reg_key:
                    with winreg.CreateKeyEx(reg_key, disallow_run_key, 0, winreg.KEY_ALL_ACCESS) as disallow_run_subkey:
                        winreg.SetValueEx(disallow_run_subkey, value_name, 0, winreg.REG_SZ, value_data)
                        logger.info(f"Set registry value for profile {profile}: {value_name} = {value_data}")
            except Exception as e:
                logger.error(f"Error setting registry value for {profile}: {e}")

def remove_registry_value_for_all_users(key, value_name):
    user_hive = winreg.HKEY_USERS
    user_profiles = []

    try:
        i = 0
        while True:
            user_profiles.append(winreg.EnumKey(user_hive, i))
            i += 1
    except OSError:
        pass  # No more subkeys

    for profile in user_profiles:
        if profile not in (".DEFAULT", "S-1-5-18", "S-1-5-19", "S-1-5-20"):
            try:
                user_key_path = f"{profile}\\{key}"
                with winreg.OpenKey(user_hive, user_key_path, 0, winreg.KEY_ALL_ACCESS) as reg_key:
                    with winreg.OpenKey(reg_key, disallow_run_key, 0, winreg.KEY_ALL_ACCESS) as disallow_run_subkey:
                        winreg.DeleteValue(disallow_run_subkey, value_name)
                        logger.info(f"Removed registry value for profile {profile}: {value_name}")
            except FileNotFoundError:
                logger.warning(f"Registry value {value_name} not found for profile {profile}")
            except Exception as e:
                logger.error(f"Error removing registry value for {profile}: {e}")

def kill_existing_instances(process_names):
    for process in psutil.process_iter(['name']):
        if process.info['name'].lower() in process_names:
            try:
                process.kill()
                logger.info(f"Killed {process.info['name']} with PID {process.pid}")
            except Exception as e:
                logger.error(f"Error killing process {process.info['name']}: {e}")

def block_apps():
    global block_print_screen
    set_registry_value_for_all_users(explorer_key_path, snipping_tool_value_name, snipping_tool_value_data)
    set_registry_value_for_all_users(explorer_key_path, steps_recorder_value_name, steps_recorder_value_data)
    kill_existing_instances([snipping_tool_value_data.lower(), steps_recorder_value_data.lower()])
    logger.info("Applications have been blocked and existing instances killed.")
    install_keyboard_hooks()
    block_print_screen = True
    logger.info("Print Screen key has been disabled.")

def unblock_apps():
    global block_print_screen
    remove_registry_value_for_all_users(explorer_key_path, snipping_tool_value_name)
    remove_registry_value_for_all_users(explorer_key_path, steps_recorder_value_name)
    logger.info("Applications have been unblocked.")
    block_print_screen = False
    uninstall_keyboard_hooks()
    logger.info("Print Screen key has been enabled.")

def prevent_new_instances():
    while True:
        try:
            window_titles = get_all_window_titles()
            block_required = False
            detected_by = None

            for hwnd, title in window_titles:
                process_name, pid = get_process_info(hwnd)
                minimized = is_window_minimized(hwnd)
                if process_name in browser_keywords:
                    for keyword in browser_keywords[process_name]:
                        if keyword.lower() in title.lower():
                            if minimized:
                                logger.info(f"Process {process_name} (PID: {pid}) with keyword {keyword} is minimized, skipping")
                                continue
                            block_required = True
                            detected_by = (process_name, pid, keyword)
                            logger.info(f"Blocking required by process {process_name} (PID: {pid}) with keyword: {keyword}")
                            break
                else:
                    for keyword in other_keywords:
                        if keyword.lower() in title.lower():
                            if minimized:
                                logger.info(f"Process {process_name} (PID: {pid}) with keyword in title {keyword} is minimized, skipping")
                                continue
                            block_required = True
                            detected_by = (process_name, pid, keyword)
                            logger.info(f"Blocking required by process {process_name} (PID: {pid}) with keyword in title: {keyword}")
                            break

            if block_required:
                if not block_print_screen:
                    block_apps()
                    logger.info(f"Blocking triggered by process {detected_by[0]} (PID: {detected_by[1]}) with keyword: {detected_by[2]}")
            else:
                if block_print_screen:
                    unblock_apps()

            # Force garbage collection to release memory
            gc.collect()

        except Exception as e:
            logger.error(f"An error occurred: {e}")

        # Check every 5 seconds
        time.sleep(5)

def main():
    logger.info("Script started.")
    prevent_new_instances()

if __name__ == "__main__":
    main()
