import os
import platform
import subprocess
import tkinter as tk
from tkinter import font, filedialog
import psutil
import winreg
from datetime import datetime
from PIL import ImageGrab
import getpass
import ctypes
import uuid
import win32com.client
import glob

def get_hostname():
    return platform.node()

def get_disk_space():
    disk = psutil.disk_usage('/')
    return f"{disk.total // (2**30)}GB"

def get_available_disk_space():
    disk = psutil.disk_usage('/')
    return f"{disk.free // (2**30)}GB"

def get_total_ram():
    ram = psutil.virtual_memory().total
    return f"{ram // (2**30)}GB"

def get_cpu():
    return platform.processor()

def check_os_compliance(os_edition, os_build, os_version):
    # Define baselines
    win11_baseline_build = "22631.3880"
    win11_baseline_version = "23H2"
    win10_baseline_build = "19045.4717"
    win10_baseline_version = "22H2"
    
    if "Windows 11" in os_edition:
        if os_build < win11_baseline_build or os_version < win11_baseline_version:
            return "FAIL", "#FF0000"
        else:
            return "PASS", "#00FF00"
    elif "Windows 10" in os_edition:
        if os_build < win10_baseline_build or os_version < win10_baseline_version:
            return "FAIL", "#FF0000"
        else:
            return "PASS", "#00FF00"
    else:
        return "Unknown OS", "yellow"

def get_os_build():
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
        build_number, _ = winreg.QueryValueEx(key, "CurrentBuildNumber")
        ubr, _ = winreg.QueryValueEx(key, "UBR")
        full_build = f"{build_number}.{ubr}"
        return full_build
    except Exception as e:
        return "Unknown"
    
def check_sendguard_files():
    files_to_check = [
        r"C:\Program Files\OCBC\SendGuard\InternalList.txt",
        r"C:\Program Files\OCBC\SendGuard\Microsoft.Office.Tools.Common.v4.0.Utilities.dll",
        r"C:\Program Files\OCBC\SendGuard\Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll",
        r"C:\Program Files\OCBC\SendGuard\SendGuard.dll",
        r"C:\Program Files\OCBC\SendGuard\SendGuard.dll.manifest",
        r"C:\Program Files\OCBC\SendGuard\SendGuard.vsto",
        r"C:\Program Files\OCBC\SendGuard\TlsList.txt"
    ]

    missing_files = [file for file in files_to_check if not os.path.exists(file)]

    if not missing_files:
        return "Found"
    else:
        return "Not Found"

def check_sendguard_service_status():
    try:
        reg_path = r"SOFTWARE\Microsoft\Office\Outlook\Addins\OCBC.SendGuard"
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
        load_behavior, _ = winreg.QueryValueEx(key, "LoadBehavior")

        if load_behavior == 3:
            return "RUNNING"
        else:
            return "NOT RUNNING"
    except FileNotFoundError:
        return "NOT RUNNING"
    except Exception as e:
        return f"Error: {str(e)}"
    
sendGuard_installation = check_sendguard_files()
sendGuard_service = check_sendguard_service_status()

def get_os_version():
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
        display_version, _ = winreg.QueryValueEx(key, "DisplayVersion")
        return display_version
    except Exception as e:
        return "Unknown"

def get_ad_ou():
    try:
        # Use PowerShell to get the DistinguishedName of the computer object
        result = subprocess.run(
            ['powershell', '-Command', '(Get-ADComputer $env:COMPUTERNAME).DistinguishedName'],
            capture_output=True, text=True
        )
        
        distinguished_name = result.stdout.strip()

        # Extract the OU from the DistinguishedName
        if distinguished_name:
            ou_parts = [part for part in distinguished_name.split(',') if part.startswith('OU=')]
            if ou_parts:
                return ', '.join(ou_parts)
        
        return "Unknown OU"
    except Exception as e:
        return f"Error: {str(e)}"

class SL_GENUINE_STATE:
    SL_GEN_STATE_IS_GENUINE = 0

def get_windows_app_id():
    try:
        wmi = win32com.client.Dispatch("WbemScripting.SWbemLocator")
        wmi_service = wmi.ConnectServer(".", "root\\cimv2")
        products = wmi_service.ExecQuery("SELECT ApplicationID FROM SoftwareLicensingProduct WHERE PartialProductKey IS NOT NULL")

        for product in products:
            app_id = product.ApplicationID
            return uuid.UUID(app_id)  # Return the Windows AppID as a UUID object

        return None
    except Exception as e:
        print(f"Error retrieving Windows AppID: {e}")
        return None

def is_genuine_windows():
    try:
        windows_app_id = get_windows_app_id()
        if windows_app_id is None:
            return "Could not retrieve Windows AppID"

        # Load the slc.dll
        slc = ctypes.WinDLL('slc.dll')

        # Define the SLIsGenuineLocal function with its signature
        SLIsGenuineLocal = slc.SLIsGenuineLocal
        SLIsGenuineLocal.argtypes = [ctypes.POINTER(ctypes.c_byte), ctypes.POINTER(ctypes.c_uint), ctypes.c_void_p]
        SLIsGenuineLocal.restype = ctypes.HRESULT

        # Prepare the arguments
        uuid_bytes = (ctypes.c_byte * 16)(*windows_app_id.bytes)
        state = ctypes.c_uint()

        # Call SLIsGenuineLocal
        hresult = SLIsGenuineLocal(uuid_bytes, ctypes.byref(state), None)

        # Check the result
        if hresult != 0:
            raise ctypes.WinError(hresult)

        return state.value == SL_GENUINE_STATE.SL_GEN_STATE_IS_GENUINE
    except Exception as e:
        return f"Error: {str(e)}"

def get_activation_state():
    try:
        if is_genuine_windows():
            return "Active"
        else:
            return "Not Activated"
    except Exception as e:
        return f"Error: {str(e)}"
    
def check_os_compliance(os_edition, os_build, os_version):
    # Define baselines
    win11_baseline_build = "22631.3880"
    win11_baseline_version = "23H2"
    win10_baseline_build = "19045.4717"
    win10_baseline_version = "22H2"

    os_build_color = "#00FF00"
    os_version_color = "#00FF00"

    if "Windows 11" in os_edition:
        if os_build < win11_baseline_build:
            os_build_color = "#FF0000"
        if os_version < win11_baseline_version:
            os_version_color = "#FF0000"
    elif "Windows 10" in os_edition:
        if os_build < win10_baseline_build:
            os_build_color = "#FF0000"
        if os_version < win10_baseline_version:
            os_version_color = "#FF0000"

    return os_build_color, os_version_color

def get_windows_version():
    return platform.version()

def get_current_user():
    return getpass.getuser()

def get_os_install_date():
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
        install_date, _ = winreg.QueryValueEx(key, "InstallDate")
        install_date = datetime.utcfromtimestamp(install_date).strftime('%Y-%m-%d %H:%M:%S')
        return install_date
    except Exception as e:
        return "Unknown"
    
def format_device_item(canvas, x, y, item_name, value, color="#00FF00", service_status="N/A", font=None):
    if font is None:
        font = ("Courier New", 11)
    
    canvas.create_text(x, y, anchor="nw", text=item_name, fill=color, font=font)  # Item name color
    canvas.create_text(x + 280, y, anchor="nw", text=value, fill=color, font=font)  # Value color
    canvas.create_line(x, y + 20, x + 800, y + 20, fill=color, width=1)  # Underline color

def format_new_teams(canvas, x, y, item_name, value, color="#00FF00", service_status="N/A", font=None):
    if font is None:
        font = ("Courier New", 11)
    
    # Column positions
    installation_col_x = 280
    service_col_x = 500
    status_col_x = 700
    
    # Draw the item name
    canvas.create_text(x, y, anchor="nw", text=item_name, fill=color, font=font)  # Item name color
    
    # Draw the installation status
    canvas.create_text(x + installation_col_x, y, anchor="nw", text=value, fill=color, font=font)  # Value color
    
    # Draw the service status
    canvas.create_text(x + service_col_x, y, anchor="nw", text=service_status, fill=color, font=font)  # Service status color
    
    # Draw the overall status (PASS/FAIL)
    status_text = "PASS" if color == "#00FF00" else "FAIL"
    canvas.create_text(x + status_col_x, y, anchor="nw", text=status_text, fill=color, font=("Courier New", 11, "bold"))
    
    # Draw the underline
    canvas.create_line(x, y + 20, x + 800, y + 20, fill=color, width=1)  # Underline color

def format_unwanted_item(canvas, x, y, item_name, value, color="#00FF00", service_status="N/A", font=None):
    if font is None:
        font = ("Courier New", 11)
    
    # Column positions
    installation_col_x = 280
    service_col_x = 500
    status_col_x = 700
    
    # Draw the item name
    canvas.create_text(x, y, anchor="nw", text=item_name, fill=color, font=font)  # Item name color
    
    # Draw the installation status
    canvas.create_text(x + installation_col_x, y, anchor="nw", text=value, fill=color, font=font)  # Value color
    
    # Draw the service status
    canvas.create_text(x + service_col_x, y, anchor="nw", text=service_status, fill=color, font=font)  # Service status color
    
    # Draw the overall status (PASS/FAIL)
    status_text = "PASS" if color == "#00FF00" else "FAIL"
    canvas.create_text(x + status_col_x, y, anchor="nw", text=status_text, fill=color, font=("Courier New", 11, "bold"))
    
    # Draw the underline
    canvas.create_line(x, y + 20, x + 800, y + 20, fill=color, width=1)  # Underline color

def check_trellix_scanner_service():
    try:
        # Get a list of all running processes
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'].lower() == 'mcshield.exe':
                return "RUNNING", "#00FF00"  # Service is running
        return "NOT RUNNING", "#FF0000"  # Service is not running
    except Exception as e:
        return f"Error: {str(e)}", "#FF0000"

trellix_scanner_status, trellix_scanner_color = check_trellix_scanner_service()

def check_cisco_anyconnect_services():
    services = ["nam", "namlm", "vpnagent"]
    service_status = {}

    for service in services:
        status, color = check_service_status(service)
        service_status[service] = (status, color)

    # Check if all requi#FF0000 services are running
    all_running = all(status == "RUNNING" for status, _ in service_status.values())

    if all_running:
        return "ALL RUNNING", "#00FF00"
    else:
        return "NOT ALL RUNNING", "#FF0000"

def check_existence(file_path, reg_path=None, expected_displayname=None, expected_version=None, should_be_present=True):
    file_exists = file_path and os.path.exists(file_path)
    registry_exists = False

    if reg_path and expected_displayname:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
                subkey_count, _, _ = winreg.QueryInfoKey(key)
                for i in range(subkey_count):
                    subkey_name = winreg.EnumKey(key, i)
                    subkey_path = f"{reg_path}\\{subkey_name}"
                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, subkey_path) as subkey:
                            display_name, _ = winreg.QueryValueEx(subkey, "DisplayName")
                            if display_name == expected_displayname:
                                if expected_version:
                                    display_version, _ = winreg.QueryValueEx(subkey, "DisplayVersion")
                                    if display_version >= expected_version:
                                        registry_exists = True
                                        break
                                else:
                                    registry_exists = True
                                    break
                    except FileNotFoundError:
                        continue
        except FileNotFoundError:
            registry_exists = False

    if should_be_present:
        if file_exists or registry_exists:
            return "Found", "#00FF00"
        else:
            return "Not Found", "#FF0000"
    else:
        if file_exists or registry_exists:
            return "Found", "#FF0000"
        else:
            return "Not Found", "#00FF00"

def get_windows_edition():
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
        product_name, _ = winreg.QueryValueEx(key, "ProductName")
        
        # Combine product name and edition with a space in between
        return f"{product_name}"
    except Exception as e:
        return "Unknown"

def check_service_status(service_name):
    try:
        result = subprocess.run(['sc', 'query', service_name], capture_output=True, text=True)
        if "STATE" in result.stdout:
            for line in result.stdout.splitlines():
                if "STATE" in line:
                    # Split on whitespace and get the last element which should be the state
                    state_value = line.split()[-1]  # This should return "RUNNING"
                    return state_value, "#00FF00" if state_value == "RUNNING" else "#FF0000"
        return "Unknown", "#FF0000"  # Return "Unknown" if the service is not found or another issue occurs
    except subprocess.CalledProcessError:
        return "Unknown", "#FF0000"
    
def check_wildcard_path(pattern):
    matching_paths = glob.glob(pattern)
    if matching_paths:
        return "Found", "#FF0000"  # Should not be present
    else:
        return "Not Found", "#00FF00"  # Expected outcome
    
def check_wildcard_path_teams(pattern):
    matching_paths = glob.glob(pattern)
    if matching_paths:
        return "Found", "#00FF00"  # Should not be present
    else:
        return "Not Found", "#FF0000"  # Expected outcome
    
def check_skype_for_business():
    try:
        reg_path = r"Software\Microsoft\Office\Outlook\Addins\UCAddin.LyncAddin.1"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path)
        load_behavior, _ = winreg.QueryValueEx(key, "LoadBehavior")
        
        if load_behavior == 3:
            return "Found", "#FF0000"  # Should not be present
        else:
            return "Not Found", "#00FF00"  # Expected outcome
    except FileNotFoundError:
        return "Not Found", "#00FF00"
    except Exception as e:
        return f"Error: {str(e)}", "#FF0000"

def get_trellix_av_dat_expeted_version():
    known_date = datetime(2024, 8, 10)
    known_value = 5614
    days_passed = (datetime.now() - known_date).days
    expected_value = known_value + days_passed
    
    return expected_value


def get_trellix_av_dat_version():
    try:
        # Define the registry path and value name
        reg_path = r"SOFTWARE\McAfee\AVSolution\DS\DS"
        value_name = "dwContentMajorVersion"

        # Open the registry key
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)

        # Retrieve the value of dwContentMajorVersion
        dat_version, _ = winreg.QueryValueEx(key, value_name)
        return dat_version

    except FileNotFoundError:
        return "Not Found"
    except Exception as e:
        return f"Error: {str(e)}"

DAT_VERSION = get_trellix_av_dat_version()
DAT_VERSION_EXPECTED = get_trellix_av_dat_expeted_version()

def format_item(canvas, x, y, item_name, existence_result, service_result, color="#00FF00"):
    # Initialize variables
    service_status = "N/A"
    service_color = existence_result[1]
    state_result = "PASS"
    state_color = "#00FF00"

    # Specific handling for Citrix Workspace
    if item_name == "Citrix Workspace":
        if existence_result[0] == "Not Found":
            color = "orange"
            state_result = "Optional"
            state_color = "orange"
            service_color = "orange"
            existence_result = ("Not Found", "orange")
        else:
            color = "#00FF00"
            state_result = "PASS"
            state_color = "#00FF00"
            service_color = "#00FF00"
            existence_result = ("Found", "#00FF00")

    # Specific handling for CyberArk EPM
    elif item_name == "CyberArk EPM":
        if existence_result[0] == "Not Found":
            color = "#FF0000"
            state_result = "FAIL"
            state_color = "#FF0000"
            service_color = "#FF0000"
            service_status = service_result[0]
            existence_result = ("Not Found", "#FF0000")
        elif service_result[0] != "RUNNING":
            color = "#00FF00"
            state_result = "PASS"
            state_color = "#00FF00"
            service_status = service_result[0]
            service_color = "orange"

    # Specific handling for Cisco AnyConnect
    elif item_name == "Cisco AnyConnect":
        if existence_result[0] == "Not Found" and service_result[0] == "NOT ALL RUNNING":
            color = "#FF0000"
            state_result = "FAIL"
            state_color = "#FF0000"
            service_color = "#FF0000"
            service_status = service_result[0]
            existence_result = ("Not Found", "#FF0000")
        elif existence_result[0] == "Found" and service_result[0] != "ALL RUNNING":
            color = "#00FF00"
            state_result = "PASS"
            state_color = "#00FF00"
            service_status = service_result[0]
            service_color = "orange"

    # Specific handling for unwanted software
    elif item_name == "Microsoft Classic Teams" or item_name == "New Outlook" or item_name == "Skype for Business" or item_name == "Solitaire" or item_name == "XBOX" or item_name == "Movie & TV":
        if existence_result[0] == "Not Found":
            color = "#00FF00"
            state_result = "PASS"
            state_color = "#00FF00"
            service_color = "#00FF00"
            existence_result = ("Not Found", "#00FF00")
        else:
            color = "#FF0000"
            state_result = "FAIL"
            state_color = "#FF0000"
            service_color = "#FF0000"

    # Specific handling for Trellix DE
    elif item_name == "Trellix Disk Encryption" and check_bitlocker()[0] == "Enabled":
        if existence_result[0] == "Not Found":
            color = "#00FF00"
            state_result = "PASS"
            state_color = "#00FF00"
            service_color = "#00FF00"
            service_status = service_result[0]
            existence_result = ("Not Found", "#00FF00")
        else:
            color = "#FF0000"
            state_result = "FAIL"
            state_color = "#FF0000"
            service_status = service_result[0]
            service_color = "#FF0000"
    elif item_name == "Trellix Disk Encryption" and check_bitlocker()[0] == "Disabled":
        if existence_result[0] == "Not Found":
            color = "#FF0000"
            state_result = "FAIL"
            state_color = "#FF0000"
            service_color = "#FF0000"
            service_status = service_result[0]
            existence_result = ("Not Found", "#FF0000")
        elif service_result[0] == "RUNNING":
            color = "#00FF00"
            state_result = "PASS"
            state_color = "#00FF00"
            service_status = service_result[0]
            service_color = "#00FF00"
            existence_result = ("Not Found", "#00FF00")

    elif item_name == "Trellix AV DAT Version":
        # Determine the status based on the difference between the actual and expected values
        try:
            if DAT_VERSION > (DAT_VERSION_EXPECTED - 28):
                color = "#00FF00"
                state_result = "PASS"
                state_color = "#00FF00"
                service_color = "#00FF00"
                service_status = "N/A"
                existence_result = (DAT_VERSION, "#00FF00")
            else:
                color = "#FF0000"
                state_result = "FAIL"
                state_color = "#FF0000"
                service_color = "#FF0000"
                service_status = "N/A"
                existence_result = (DAT_VERSION, "#FF0000")
        except Exception as e:
            color = "#FF0000"
            state_result = "FAIL"
            state_color = "#FF0000"
            service_status = "N/A"
            service_color = "#FF0000"
    elif item_name == "Trellix Scanner Service":
        # Determine the status based on the difference between the actual and expected values
        if trellix_scanner_status == "RUNNING":
            color = "#00FF00"
            state_result = "PASS"
            state_color = "#00FF00"
            service_color = "#00FF00"
            service_status = "RUNNING"
            existence_result = ("Found", "#00FF00")
        else:
            color = "#FF0000"
            state_result = "FAIL"
            state_color = "#FF0000"
            service_color = "#FF0000"
            service_status = "NOT RUNNING"
            existence_result = (DAT_VERSION, "#FF0000")
    elif item_name == "SendGuard":
        # Determine the status based on the difference between the actual and expected values
        if sendGuard_installation == "Found" and sendGuard_service == "RUNNING":
            color = "#00FF00"
            state_result = "PASS"
            state_color = "#00FF00"
            service_color = "#00FF00"
            service_status = "RUNNING"
            existence_result = ("Found", "#00FF00")
        elif sendGuard_installation == "Found" and sendGuard_service == "NOT RUNNING":
            color = "#FF0000"
            state_result = "FAIL"
            state_color = "#FF0000"
            service_color = "#FF0000"
            service_status = "NOT RUNNING"
            existence_result = ("Found", "#00FF00")
        else:
            color = "#FF0000"
            state_result = "FAIL"
            state_color = "#FF0000"
            service_color = "#FF0000"
            service_status = "NOT RUNNING"
            existence_result = ("Not Found", "#FF0000")

    else:
        # Determine the appropriate service result
        if isinstance(service_result, tuple):
            service_status, service_color = service_result
        else:
            service_status, service_color = ("N/A", existence_result[1])

        # Determine the final state result
        if service_status == "N/A":
            state_result = "PASS" if existence_result[0] == "Found" else "FAIL"
            state_color = existence_result[1]
        else:
            state_result = "PASS" if existence_result[1] == "#00FF00" and service_status == "RUNNING" else "FAIL"
            state_color = "#00FF00" if state_result == "PASS" else "#FF0000"

    # Render the item details
    canvas.create_text(x, y, anchor="nw", text=f"{item_name}", fill=color, font=("Courier New", 11))
    canvas.create_text(x + 280, y, anchor="nw", text=existence_result[0], fill=color, font=("Courier New", 11))
    canvas.create_text(x + 500, y, anchor="nw", text=service_status, fill=service_color, font=("Courier New", 11))
    canvas.create_text(x + 700, y, anchor="nw", text=state_result, fill=state_color, font=("Courier New", 11, "bold"))

    # Draw the underline
    canvas.create_line(x, y + 20, x + 800, y + 20, fill=color, width=1)


def save_as_image(canvas):
    file_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png"), ("All files", "*.*")])
    if file_path:
        x = root.winfo_rootx() + canvas.winfo_x()
        y = root.winfo_rooty() + canvas.winfo_y()
        x1 = x + canvas.winfo_width()
        y1 = y + canvas.winfo_height()
        ImageGrab.grab().crop((x, y, x1, y1)).save(file_path)

def draw_category(canvas, category_name, items, start_y):
    y = start_y
    height = 30 * (len(items) + 1)
    canvas.create_text(20, y, anchor="nw", text=category_name, font=("Courier New", 12, "bold", "underline"), fill="yellow")
    y += 30
    return y, height

def check_bitlocker():
    try:
        result = subprocess.run(['manage-bde', '-status', 'C:'], capture_output=True, text=True)
        if "Protection On" in result.stdout:
            return "Enabled", "#00FF00"
        else:
            return "Disabled", "#FF0000"
    except Exception as e:
        return "Unknown", "#FF0000"

def check_trellix_disk_encryption():
    try:
        key_path = r"SOFTWARE\McAfee\Endpoint Encryption Agent"
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path)
        version, _ = winreg.QueryValueEx(key, "Version")
        encryption_state, _ = winreg.QueryValueEx(key, "EncryptionState")
        if version:
            if encryption_state == 1:  # Assuming 1 means enabled
                return "Found/Enabled", "#00FF00"
            else:
                return "Found/Disabled", "#FF0000"
        else:
            return "Not Found", "#FF0000"
    except FileNotFoundError:
        return "Not Found", "#FF0000"
    except Exception as e:
        return "Unknown", "#FF0000"

def main():
    global root
    root = tk.Tk()
    root.title("Compliance Check Report")

    # Calculate the requi#FF0000 width and set height to 100 pixels
    canvas_width = 850  # Adjusted width to fit the longest row
    canvas_height = 500  # Set the height to 100 pixels
    margin = 20  # Define margin before using it

    # Make the window's width unchangeable
    root.resizable(width=False, height=False)

    menu_bar = tk.Menu(root)
    file_menu = tk.Menu(menu_bar, tearoff=0)
    file_menu.add_command(label="Save as Image", command=lambda: save_as_image(canvas))
    file_menu.add_separator()
    file_menu.add_command(label="Exit", command=root.quit)
    menu_bar.add_cascade(label="File", menu=file_menu)
    root.config(menu=menu_bar)

    # Set up a scrollable canvas
    frame = tk.Frame(root)
    frame.pack(fill="both", expand=True)

    canvas = tk.Canvas(frame, width=canvas_width, height=canvas_height, bg="black")
    canvas.pack(side="left", fill="both", expand=True)

    scrollbar = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="right", fill="y")

    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    # Set up fonts
    header_font = font.Font(family="Courier New", size=14, weight="bold")
    subheader_font = font.Font(family="Courier New", size=12, weight="bold")
    normal_font = font.Font(family="Courier New", size=11)
    bold_font = font.Font(family="Courier New", size=11, weight="bold")

    y = 20

    # Header Information
    today_date = datetime.now().strftime("%d %B %Y")
    date_x = canvas_width - margin - 150  # Adjusted value to align the date with a margin

    canvas.create_text(margin, y, anchor="nw", text="Compliance Check Report", font=header_font, fill="white")
    canvas.create_text(date_x, y, anchor="nw", text=today_date, font=header_font, fill="white")
    y += 30

    # System Info
    hostname = get_hostname()
    current_user = get_current_user()
    disk_space = get_disk_space()
    available_disk_space = get_available_disk_space()
    total_ram = get_total_ram()
    cpu = get_cpu()
    windows_edition = get_windows_edition()
    os_install_date = get_os_install_date()

    system_info = [
        f"Hostname: {hostname}",
        f"Current Logon User ID: {current_user}",
        f"Total Disk Space: {disk_space}",
        f"Available Disk Space: {available_disk_space}",
        f"Total RAM: {total_ram}",
        f"CPU: {cpu}",
        f"Windows OS Edition: {windows_edition}",
        f"OS Installation Date: {os_install_date}"
    ]

    for info in system_info:
        canvas.create_text(margin, y, anchor="nw", text=info, font=normal_font, fill="white")
        y += 20

    y += 20

    # Device Information Category
    os_build = get_os_build()
    os_version = get_os_version()
    activation_state = get_activation_state()
    ad_ou = get_ad_ou()

    os_build_color, os_version_color = check_os_compliance(windows_edition, os_build, os_version)
    activation_color = "#00FF00" if activation_state == "Active" else "#FF0000"
    ad_ou_color = "#00FF00" if ad_ou != "Unknown" else "#FF0000"

    # Determine overall compliance
    if (os_build_color == "#FF0000" or os_version_color == "#FF0000" or 
        activation_color == "#FF0000" or ad_ou_color == "#FF0000"):
        compliance_status = "FAIL"
        compliance_color = "#FF0000"
    else:
        compliance_status = "PASS"
        compliance_color = "#00FF00"

    device_info_items = [
        {"name": "OS Build", "value": os_build, "color": os_build_color},
        {"name": "OS Version", "value": os_version, "color": os_version_color},
        {"name": "AD OU", "value": ad_ou, "color": ad_ou_color},
        {"name": "Windows Activation State", "value": activation_state, "color": activation_color},
        {"name": "Compliance Status", "value": compliance_status, "color": compliance_color, "font": bold_font}
    ]

    y, _ = draw_category(canvas, "Device Information", device_info_items, y)
    for item in device_info_items:
        format_device_item(canvas, 20, y, item["name"], item["value"], item.get("color", "#00FF00"), item.get("font", normal_font))
        y += 30

    y += 20

    # Column Titles
    canvas.create_text(20, y, anchor="nw", text="Item", font=subheader_font, fill="white")
    canvas.create_text(300, y, anchor="nw", text="Installation", font=subheader_font, fill="white")
    canvas.create_text(520, y, anchor="nw", text="Service", font=subheader_font, fill="white")
    canvas.create_text(720, y, anchor="nw", text="Status", font=subheader_font, fill="white")
    y += 30

    # Define the categories and items
    categories = {
        "Security": [
            {
                "name": "Trellix Agent",
                "file_path": r"C:\Program Files\McAfee\Agent\cmdagent.exe",
                "reg_path": None,
                "service_name": "masvc"
            },
            {
                "name": "Trellix AV DAT Version",
                "file_path": None,
                "reg_path": r"SOFTWARE\McAfee\AVSolution\DS\DS",
                "service_name": None
            },
            {
                "name": "Eclypsium",
                "file_path": r"C:\Program Files\Eclypsium\EclypsiumApp.exe",
                "reg_path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                "expected_displayname": "Eclypsium",
                "expected_version": "3.4.00",
                "service_name": "EclypsiumService"
            },
            {
                "name": "Carbon Black",
                "file_path": r"C:\Program Files (x86)\CarbonBlack",
                "reg_path": r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
                "expected_displayname": "VMWare Carbon Black EDR Sensor",
                "expected_version": "7.4.1.18957",
                "service_name": "CarbonBlack"
            },
            {
                "name": "BitLocker",
                "file_path": None,  # Not needed since installation status is checked by manage-bde
                "reg_path": None,   # Not needed
                "service_name": "BDESVC"
            },
            {
                "name": "Trellix Disk Encryption",
                "file_path": None,
                "reg_path": None,
                "service_name": "MfeFfCoreService"
            },
            {
                "name": "Trellix Scanner Service",
                "file_path": r"C:\Program Files\Common Files\McAfee\AVSolution\mcshield.exe",
                "reg_path": None,
                "service_name": "mcshield"
            },
            {
                "name": "CyberArk EPM",
                "file_path": r"C:\Program Files\CyberArk\Endpoint Privilege manager\Agent\vf_agent.exe",
                "reg_path": None,
                "service_name": "vf_agent"
            }
        ],
        "DLP": [
            {
                "name": "Trellix DLP",
                "file_path": r"C:\Program Files\McAfee\DLP\Agent\fcag.exe",
                "reg_path": None,
                "service_name": "TrellixDLPAgentService"
            },
            {
                "name": "SendGuard",
                "file_path": r"C:\Program Files\OCBC\SendGuard\SendGuard.dll",
                "reg_path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{86BA39A1-BBCD-4ADB-B150-38E748422BD8}",
                "expected_displayname": None,
                "expected_version": "1.0.9",
                "service_name": None
            },
            {
                "name": "AgileMark",
                "file_path": r"C:\Program Files (x86)\AgileMark\AgileMark.exe",
                "reg_path": r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
                "expected_displayname": "AgileMark",
                "expected_version": "1.0.13.6",
                "service_name": "AgileService.exe"
            }
        ],
        "Standard Software": [
            {
                "name": "Citrix Workspace",
                "file_path": r"C:\Program Files (x86)\Citrix\ICA Client\wfcrun32.exe",
                "reg_path": None,
                "service_name": None,
                "Optional": True
            },
            {
                "name": "Cisco AnyConnect",
                "file_path": r"C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnagent.exe",
                "reg_path": None,
                "service_name": None  # Exception, will handle this separately
            },
            {
                "name": "Google Chrome",
                "file_path": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                "reg_path": None,
                "service_name": None
            },
            {
                "name": "7-zip",
                "file_path": r"C:\Program Files\7-Zip\7z.exe",
                "reg_path": None,
                "service_name": None
            },
            {
                "name": "Microsoft Edge",
                "file_path": r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                "reg_path": None,
                "service_name": None
            },
            {
                "name": "Office 365",
                "file_path": r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
                "reg_path": None,
                "service_name": None
            },
            {
                "name": "Microsoft Teams",
                "check_function": lambda: check_wildcard_path_teams(r"C:\Program Files\WindowsApps\*MSTeams*"),
                "description": "Check if New Teams is installed (should not be present).",
                "service_name": None
            }
        ]
    }

    checks = {
        "Unwanted Software": [
            {
                "name": "New Outlook",
                "check_function": lambda: check_wildcard_path(r"C:\Program Files\WindowsApps\*OutlookForWindows*"),
                "description": "Check if New Outlook is installed (should not be present).",
                "service_name": None
            },
            {
                "name": "XBOX",
                "check_function": lambda: check_wildcard_path(r"C:\Program Files\WindowsApps\*xboxapp*"),
                "description": "Check if XBOX is installed (should not be present)."
            },
            {
                "name": "Classic Teams",
                "check_function": lambda: check_wildcard_path(r"C:\Program Files\WindowsApps\MicrosoftTeams*"),
                "description": "Check if Classic Teams is installed (should not be present)."
            },
            {
                "name": "Solitaire",
                "check_function": lambda: check_wildcard_path(r"C:\Program Files\WindowsApps\*MicrosoftSolitaireCollection*"),
                "description": "Check if Solitaire is installed (should not be present)."
            },
            {
                "name": "Movie & TV",
                "check_function": lambda: check_wildcard_path(r"C:\Program Files\WindowsApps\*ZuneVideo*"),
                "description": "Check if Movie & TV is installed (should not be present)."
            },
            {
                "name": "Skype for Business",
                "check_function": check_skype_for_business,
                "description": "Check if Skype for Business is installed (should not be present)."
            }
        ]
    }

    for category_name, items in categories.items():
        items = sorted(items, key=lambda x: x["name"])

        y, _ = draw_category(canvas, category_name, items, y)
        for item in items:
            if item["name"] == "Cisco AnyConnect":
                existence_result = check_existence(item["file_path"], item.get("reg_path"), item.get("expected_displayname"), item.get("expected_version"))
                service_result, service_color = check_cisco_anyconnect_services()
                format_item(canvas, 20, y, item["name"], existence_result, (service_result, service_color), existence_result[1])
            elif item["name"] == "BitLocker":
                existence_result = check_bitlocker()  # Get the BitLocker status directly
                service_result, service_color = check_service_status(item["service_name"])
                format_item(canvas, 20, y, item["name"], existence_result, (service_result, service_color), existence_result[1])
            elif item["name"] == "Microsoft Teams":
                name = item["name"]
                check_function = item["check_function"]
                result, color = check_function()
                format_new_teams(canvas, 20, y, name, result, color)
            else:
                existence_result = check_existence(item["file_path"], item.get("reg_path"), item.get("expected_displayname"), item.get("expected_version"))
                service_result, service_color = check_service_status(item["service_name"]) if item["service_name"] else ("N/A", existence_result[1])
                format_item(canvas, 20, y, item["name"], existence_result, (service_result, service_color), existence_result[1])
            y += 30

    # Running the checks
    for category, items in checks.items():

        y, _ = draw_category(canvas, category, items, y)
        for item in items:
            name = item["name"]
            check_function = item["check_function"]
            
            # Run the check
            result, color = check_function()
            
            # Display result on the canvas
            format_unwanted_item(canvas, 20, y, name, result, color)
            y += 30

    root.mainloop()

if __name__ == "__main__":
    main()
