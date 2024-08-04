# To be used with examine_known_executables.py

import csv
import time
import psutil
import logging
import logging.handlers
import ctypes
from ctypes import wintypes
import gc
import os
import signal
import sys
import pefile
import hashlib
from time import sleep

# Define the paths
csv_file_path = r"C:\Program Files\OCBC\OCBCDLP\BlockedApps.csv"
log_file_dir = r"C:\temp\PCEng\OCBCDLP"
pre_existing_logs_dir = r"C:\Program Files\OCBC\OCBCDLP"
log_file_path = os.path.join(log_file_dir, "BlockedApps.log")

# Create the log directory if it doesn't exist
os.makedirs(log_file_dir, exist_ok=True)

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

OpenProcess = kernel32.OpenProcess
OpenProcessToken = advapi32.OpenProcessToken
GetTokenInformation = advapi32.GetTokenInformation
CloseHandle = kernel32.CloseHandle

whitelisted_processes = set()
checked_processes = set()

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
                row['Hash'] = calculate_file_hash(row['FilePath'])
                blocked_apps.append(row)
    return blocked_apps

def calculate_file_hash(file_path):
    """Calculate the SHA-256 hash of a file."""
    sha256 = hashlib.sha256()
    try:
        with open(file_path, 'rb') as f:
            while chunk := f.read(8192):
                sha256.update(chunk)
    except Exception as e:
        logging.error(f"Error calculating hash for {file_path}: {e}")
        return None
    return sha256.hexdigest()

def get_import_table(file_path):
    """Get the import table of an executable."""
    try:
        pe = pefile.PE(file_path)
        if not hasattr(pe, 'DIRECTORY_ENTRY_IMPORT'):
            return []

        imports = []
        for entry in pe.DIRECTORY_ENTRY_IMPORT:
            for imp in entry.imports:
                if imp.name:
                    imports.append(f"{entry.dll.decode('utf-8')}.{imp.name.decode('utf-8')}")
                else:
                    imports.append(f"{entry.dll.decode('utf-8')}.{imp.ordinal}")
        return imports
    except Exception as e:
        logging.error(f"Error retrieving import table for {file_path}: {e}")
        return []

def compare_import_tables(imports1, imports2):
    """Compare two import tables allowing a maximum of 1 difference."""
    set1 = set(imports1)
    set2 = set(imports2)
    differences = len(set1.symmetric_difference(set2))
    return differences <= 1

def compare_data_directories(dir1, dir2):
    """Compare two data directories allowing a maximum of 1 difference."""
    differences = 0
    if dir1.VirtualAddress != dir2.VirtualAddress:
        differences += 1
    if dir1.Size != dir2.Size:
        differences += 1
    return differences <= 1

def compare_optional_headers_and_data_directories(pe1, pe2):
    """Compare the optional headers and data directories of two PE files allowing a maximum of 1 difference."""
    header1 = pe1.OPTIONAL_HEADER
    header2 = pe2.OPTIONAL_HEADER
    data_directories1 = pe1.OPTIONAL_HEADER.DATA_DIRECTORY
    data_directories2 = pe2.OPTIONAL_HEADER.DATA_DIRECTORY

    differences = 0
    header_fields = [
        'Magic', 'MajorLinkerVersion', 'MinorLinkerVersion', 'SizeOfCode', 'SizeOfInitializedData',
        'SizeOfUninitializedData', 'AddressOfEntryPoint', 'BaseOfCode', 'ImageBase', 'SectionAlignment',
        'FileAlignment', 'MajorOperatingSystemVersion', 'MinorOperatingSystemVersion', 'MajorImageVersion',
        'MinorImageVersion', 'MajorSubsystemVersion', 'MinorSubsystemVersion', 'SizeOfImage', 'SizeOfHeaders',
        'CheckSum', 'Subsystem', 'DllCharacteristics', 'SizeOfStackReserve', 'SizeOfStackCommit',
        'SizeOfHeapReserve', 'SizeOfHeapCommit', 'LoaderFlags', 'NumberOfRvaAndSizes'
    ]

    for field in header_fields:
        if getattr(header1, field) != getattr(header2, field):
            differences += 1
            if differences > 1:
                return False

    for dir1, dir2 in zip(data_directories1, data_directories2):
        if not compare_data_directories(dir1, dir2):
            differences += 1
            if differences > 1:
                return False

    return differences <= 1

def capture_unique_dlls(proc):
    """Capture the unique DLLs accessed by a running process within 2 seconds."""
    try:
        unique_dlls = set()
        start_time = time.time()
        
        while time.time() - start_time < 2:
            for mem_map in proc.memory_maps(grouped=False):
                if mem_map.path and mem_map.path.lower().startswith(r'c:\windows') and mem_map.path.lower().endswith('.dll'):
                    unique_dlls.add(mem_map.path)
        
        return unique_dlls
    except Exception as e:
        logging.error(f"Error capturing unique DLLs for process {proc.pid}: {e}")
        return set()

def compare_unique_dlls(logged_dlls, current_dlls):
    """Compare unique DLLs for 80% similarity."""
    if len(logged_dlls) == 0 or len(current_dlls) == 0:
        return False, 0

    # Use the larger set as the base set
    base_dlls = logged_dlls if len(logged_dlls) > len(current_dlls) else current_dlls
    smaller_set = current_dlls if len(logged_dlls) > len(current_dlls) else logged_dlls

    match_count = len(base_dlls & smaller_set)
    match_percentage = (match_count / len(base_dlls)) * 100

    return match_percentage >= 80, match_percentage

def match_process(proc, blocked_apps):
    """Match a process against the list of blocked apps."""
    try:
        proc_exe = proc.exe()
        if not proc_exe or any(proc_exe.lower().startswith(path) for path in [
            r'c:\windows', r'c:\program files', r'c:\program files (x86)', r'c:\ocbc'
        ]):
            return False, {}, 0, "", {}

        # Get current DLLs
        current_dlls = capture_unique_dlls(proc)

        # Get import table and optional headers + data directories of the current process
        proc_imports = get_import_table(proc_exe)
        proc_hash = calculate_file_hash(proc_exe)
        pe_proc = pefile.PE(proc_exe)  # Load the process PE file

        match_results = {
            "DLLs": False,
            "Import Table": False,
            "Optional Headers and Data Directories": False
        }
        match_percentage = 0
        for blocked_app in blocked_apps:
            blocked_app_exe = blocked_app['FilePath']
            blocked_app_log_file = os.path.join(pre_existing_logs_dir, f"{os.path.basename(blocked_app_exe)}_libraries_log.txt")

            if not os.path.exists(blocked_app_log_file):
                logging.warning(f"Log file does not exist for blocked app: {blocked_app_exe}")
                continue

            # Compare unique DLLs
            with open(blocked_app_log_file, 'r') as f:
                logged_dlls = set(line.strip() for line in f.readlines() if line.strip().endswith('.dll'))

            is_match_dlls, match_percentage_dlls = compare_unique_dlls(logged_dlls, current_dlls)
            match_results["DLLs"] = is_match_dlls
            if is_match_dlls:
                match_percentage = match_percentage_dlls

            # Compare import table
            blocked_app_imports = get_import_table(blocked_app_exe)
            is_match_import_table = compare_import_tables(proc_imports, blocked_app_imports)
            match_results["Import Table"] = is_match_import_table

            # Compare optional headers + data directories
            pe_blocked = pefile.PE(blocked_app_exe)
            is_match_optional_headers = compare_optional_headers_and_data_directories(pe_proc, pe_blocked)
            match_results["Optional Headers and Data Directories"] = is_match_optional_headers

            if any(match_results.values()):
                return True, blocked_app, match_percentage, match_results

        return False, {}, 0, match_results

    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
        logging.error(f"Error accessing process details: {e}")
    except Exception as e:
        logging.error(f"General error: {e}")
    return False, {}, 0, {"DLLs": False, "Import Table": False, "Optional Headers and Data Directories": False}

def terminate_matching_processes(blocked_apps):
    """Terminate matching processes and log the actions."""
    for proc in psutil.process_iter(['pid', 'exe']):
        try:
            if proc.pid in whitelisted_processes or proc.pid in checked_processes:
                continue

            proc_exe = proc.info['exe']
            if not proc_exe or any(proc_exe.lower().startswith(path) for path in [
                r'c:\windows', r'c:\program files', r'c:\program files (x86)', r'c:\ocbc'
            ]):
                continue

            logging.info(f"Checking process: {proc_exe} (PID: {proc.pid})")
            is_match, blocked_app, match_percentage, match_results = match_process(proc, blocked_apps)
            if is_match:
                if not is_admin_process(proc):
                    proc.terminate()
                    log_message = (
                        f"Terminated process: {proc_exe} | "
                        f"Blocked App: {blocked_app} | "
                        f"Similarity: {match_percentage}% | "
                        f"Match Results: {match_results}"
                    )
                    logging.info(log_message)
                else:
                    log_message = (
                        f"Matched but not terminated (admin/system): {proc_exe} | "
                        f"Blocked App: {blocked_app} | "
                        f"Similarity: {match_percentage}% | "
                        f"Match Results: {match_results}"
                    )
                    logging.info(log_message)
            else:
                logging.info(f"No match found for process: {proc_exe} (PID: {proc.pid})")
            checked_processes.add(proc.pid)

        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
            logging.error(f"Error processing: {e}")
    # Release resources
    gc.collect()

def whitelist_existing_processes():
    """Whitelist all existing processes at the time the script starts."""
    for proc in psutil.process_iter(['pid', 'exe']):
        try:
            whitelisted_processes.add(proc.pid)
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue

def main():
    """Main function to run the script."""
    whitelist_existing_processes()
    blocked_apps = load_blocked_apps(csv_file_path)
    # Check all existing processes at startup
    terminate_matching_processes(blocked_apps)
    while True:
        terminate_matching_processes(blocked_apps)
        sleep(5)  # Check every 5 seconds
        # Release resources
        gc.collect()

if __name__ == "__main__":
    def signal_handler(sig, frame):
        sys.exit(0)
    signal.signal(signal.SIGINT, signal_handler)
    main()
