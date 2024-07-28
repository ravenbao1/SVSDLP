import csv
import psutil
import time
import logging
import logging.handlers
import ctypes
from ctypes import wintypes
import gc

# Define the path to the CSV file and log file
csv_file_path = r"C:\Program Files\OCBC\OCBCDLP\BlockedApps.csv"
log_file_path = r"C:\Program Files\OCBC\OCBCDLP\BlockedApps.log"

# Set up logging with rotation
log_handler = logging.handlers.RotatingFileHandler(
    log_file_path, maxBytes=5 * 1024 * 1024, backupCount=5)
logging.basicConfig(handlers=[log_handler], level=logging.INFO, format='%(asctime)s - %(message)s')

# Log the start of the script
logging.info("Script started.")

# Constants for accessing file properties
kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
advapi32 = ctypes.WinDLL('advapi32', use_last_error=True)

PROCESS_QUERY_INFORMATION = 0x0400
PROCESS_VM_READ = 0x0010
TOKEN_QUERY = 0x0008
TokenElevation = 20

GetFileVersionInfoSize = ctypes.windll.version.GetFileVersionInfoSizeW
GetFileVersionInfo = ctypes.windll.version.GetFileVersionInfoW
VerQueryValue = ctypes.windll.version.VerQueryValueW
OpenProcess = kernel32.OpenProcess
OpenProcessToken = advapi32.OpenProcessToken
GetTokenInformation = advapi32.GetTokenInformation
CloseHandle = kernel32.CloseHandle

def is_admin_process(proc):
    """Check if the process is running with admin privileges."""
    try:
        process_handle = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, False, proc.pid)
        if not process_handle:
            return False
        token_handle = wintypes.HANDLE()
        if not OpenProcessToken(process_handle, TOKEN_QUERY, ctypes.byref(token_handle)):
            CloseHandle(process_handle)
            return False
        elevation = wintypes.DWORD()
        size = wintypes.DWORD(ctypes.sizeof(elevation))
        if not GetTokenInformation(token_handle, TokenElevation, ctypes.byref(elevation), size, ctypes.byref(size)):
            CloseHandle(token_handle)
            CloseHandle(process_handle)
            return False
        is_admin = elevation.value == 1
        CloseHandle(token_handle)
        CloseHandle(process_handle)
        return is_admin
    except Exception as e:
        logging.error(f"Error checking if process is admin: {e}")
        return False

def load_blocked_apps(csv_file_path):
    """Load the list of blocked apps from the CSV file."""
    blocked_apps = []
    with open(csv_file_path, mode='r') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            if all(row[key] for key in row):  # Skip rows with missing values
                blocked_apps.append(row)
    return blocked_apps

def get_file_version_info(file_path):
    """Get the version info of a file."""
    size = GetFileVersionInfoSize(file_path, None)
    if size == 0:
        return None

    res = ctypes.create_string_buffer(size)
    if not GetFileVersionInfo(file_path, 0, size, res):
        return None

    return res

def query_value(buffer, sub_block):
    """Query a specific value from the file version info."""
    r = wintypes.LPVOID()
    l = wintypes.UINT()
    if not VerQueryValue(buffer, sub_block, ctypes.byref(r), ctypes.byref(l)):
        return None
    return ctypes.wstring_at(r, l.value).strip('\x00')

def get_file_properties(file_path, keys):
    """Get the file properties for comparison."""
    props = {key: None for key in keys}
    
    try:
        buffer = get_file_version_info(file_path)
        if buffer:
            for key in keys:
                if key == 'Type':
                    continue  # Skip querying 'Type' as it's not in version info
                value = query_value(buffer, f'\\StringFileInfo\\040904b0\\{key}')
                if key == 'OriginalFilename' and value and value.lower().endswith('.mui'):
                    value = value[:-4]  # Remove .mui extension
                props[key] = value if value else None

        # Manually set the 'Type' property if it exists in the keys
        if 'Type' in keys:
            props['Type'] = 'Application' if file_path.lower().endswith('.exe') else 'DLL'
    except Exception as e:
        logging.error(f"Error retrieving file properties for {file_path}: {e}")
        return None
    
    return props

def match_process(proc, blocked_apps):
    """Match a process against the list of blocked apps."""
    try:
        proc_exe = proc.info['exe']
        if not proc_exe:
            return False, {}, {}

        for blocked_app in blocked_apps:
            keys = blocked_app.keys()
            file_props = get_file_properties(proc_exe, keys)
            if file_props is None:
                continue

            all_match = True
            for key in keys:
                app_value = blocked_app[key]
                prop_value = file_props[key]
                if not (app_value == '*' or app_value.lower() == (prop_value.lower() if prop_value else '')):
                    all_match = False
                    break

            if all_match:
                return True, file_props, blocked_app
        
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
        logging.error(f"Error accessing process details: {e}")
    return False, {}, {}

def terminate_matching_processes(blocked_apps):
    """Terminate matching processes and log the actions."""
    for proc in psutil.process_iter(['pid', 'exe']):
        try:
            proc_exe = proc.info['exe']
            if proc_exe:
                is_match, file_props, blocked_app = match_process(proc, blocked_apps)
                if is_match:
                    if not is_admin_process(proc):
                        proc.terminate()
                        log_message = (
                            f"Terminated process: {proc_exe} | "
                            f"Process properties: {file_props} | "
                            f"Blocked App: {blocked_app}"
                        )
                        logging.info(log_message)
                    else:
                        log_message = (
                            f"Matched but not terminated (admin/system): {proc_exe} | "
                            f"Process properties: {file_props} | "
                            f"Blocked App: {blocked_app}"
                        )
                        logging.info(log_message)
                    
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
            logging.error(f"Error processing: {e}")
    # Release resources
    gc.collect()

def main():
    """Main function to run the script."""
    blocked_apps = load_blocked_apps(csv_file_path)
    while True:
        terminate_matching_processes(blocked_apps)
        time.sleep(5)  # Check every 5 seconds
        # Release resources
        gc.collect()

if __name__ == "__main__":
    main()
