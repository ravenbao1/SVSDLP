import csv
import time
import psutil
import os
import signal
import sys

def log_application_behavior(app_path, log_file):
    try:
        # Start the application
        process = psutil.Popen(app_path)
        
        # Capture unique DLLs accessed by the process under C:\Windows\*
        unique_dlls = set()
        start_time = time.time()
        
        while time.time() - start_time < 2:
            for mem_map in process.memory_maps(grouped=False):
                if mem_map.path and mem_map.path.lower().startswith(r'c:\windows') and mem_map.path.lower().endswith('.dll'):
                    unique_dlls.add(mem_map.path)
        
        # Terminate the process
        process.terminate()
        process.wait()

        # Write the unique DLLs to a log file without sorting
        with open(log_file, 'w') as f:
            for dll in unique_dlls:
                f.write(f"{dll}\n")

    except Exception as e:
        print(f"An error occurred while processing {app_path}: {e}")

def main():
    csv_file_path = r'C:\Program Files\OCBC\OCBCDLP\BlockedApps.csv'
    log_directory = r'C:\Program Files\OCBC\OCBCDLP'
    
    if not os.path.exists(csv_file_path):
        print(f"The file {csv_file_path} does not exist.")
        return
    
    if not os.path.exists(log_directory):
        os.makedirs(log_directory)
    
    with open(csv_file_path, 'r') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            app_path = row['FilePath']
            log_file = os.path.join(log_directory, f"{os.path.basename(app_path)}_libraries_log.txt")
            log_application_behavior(app_path, log_file)

if __name__ == "__main__":
    def signal_handler(sig, frame):
        sys.exit(0)
    signal.signal(signal.SIGINT, signal_handler)
    main()
