import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import sqlite3
import threading
import time
import numpy as np
import traceback
import requests
import urllib3
import os
from msal import PublicClientApplication

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class DeviceInventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Device Inventory Viewer")       

        # Proxy configuration
        self.PROXY_URL = "http://gosurf.ocbc.local:8080"  # Replace with your proxy URL
        self.PROXIES = {
            "http": self.PROXY_URL,
            "https": self.PROXY_URL,
        }
        self.session = requests.Session()
        self.session.proxies.update(self.PROXIES)

        # Attempt to test the proxy
        if not self.test_proxy():
            # If proxy is not reachable, use direct traffic
            self.PROXIES = None
            self.session.proxies.clear()

        # Intune Graph API credentials

        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.scope = ["User.Read"]

        # Initialize SQLite database and GUI components after authentication
        self.initialize_main_application()        

        # Bind the configure event to track window state changes
        self.root.bind("<Configure>", self.on_window_configure)

    def test_proxy(self):
        """Test if the proxy is reachable."""
        try:
            # Use a lightweight request to test the proxy
            test_url = "https://www.google.com"
            response = self.session.get(test_url, verify=False, proxies=self.PROXIES, timeout=5)
            if response.status_code == 200:
                print("Proxy server is reachable.")
                return True
        except requests.exceptions.RequestException as e:
            print(f"Proxy test failed: {e}. Falling back to direct traffic.")
        return False

    def initialize_main_application(self):
        # Initialize SQLite database
        self.conn = sqlite3.connect('inventory.db', timeout=30, check_same_thread=False)
        self.db_lock = threading.Lock()
        self.cursor = self.conn.cursor()
        self.conn_thread_id = threading.get_ident()
        self.create_tables()   
        # Set page size for pagination
        self.page_size = 100
        self.software_page_size = 100
        self.current_page = 0
        # Continue with your application setup, such as connecting to DB and building the GUI
        self.set_window_size()
        self.selection_window()

    def create_tables(self):
        # Create Devices, Remarks, and BitLockerKeys tables if not exist
        self.execute_query('''CREATE TABLE IF NOT EXISTS Devices (
                                DeviceName TEXT,
                                SerialNumber TEXT,
                                EntraDeviceID TEXT,
                                IntuneDeviceID TEXT,
                                UserPrincipalName TEXT,
                                OperatingSystem TEXT,
                                OSVersion TEXT,
                                ComplianceState TEXT,
                                Model TEXT,
                                Manufacturer TEXT,
                                MAC TEXT,
                                IntuneLastSync TEXT,
                                ReportTime TEXT,
                                Source TEXT,
                                Encryption TEXT,
                                UserDisplayName TEXT,
                                JobTitle TEXT,
                                Department TEXT,
                                City TEXT,
                                Country TEXT,
                                TrustType TEXT,
                                TotalStorage TEXT,
                                FreeStorage TEXT,
                                PhysicalMemory TEXT,
                                PRIMARY KEY (SerialNumber, DeviceName)
                              )''')

        self.execute_query('''CREATE TABLE IF NOT EXISTS Remarks (
                                SerialNumber TEXT,
                                DeviceName TEXT,
                                Remarks TEXT,
                                PRIMARY KEY (SerialNumber, DeviceName)
                              )''')
        
        self.execute_query('''CREATE TABLE IF NOT EXISTS BitLockerKeys (
                                DeviceName TEXT,
                                SerialNumber TEXT,
                                KeyID TEXT,
                                RecoveryKey TEXT,
                                BackupTime TEXT,
                                PRIMARY KEY (SerialNumber, DeviceName)
                              )''')
        
        self.execute_query('''CREATE TABLE IF NOT EXISTS Software (
                            SoftwareName TEXT,
                            Version TEXT,
                            InstalledDevices INTEGER,
                            PRIMARY KEY (SoftwareName, Version)
                          )''')

        with self.db_lock:
            self.conn.commit()

    def set_window_size(self):
        self.root.minsize(width=400, height=600)
        self.root.geometry(f"400x600")
        
    """ def authenticate_and_check_access(self):
        try:
            # Configuration for MSAL and Microsoft Graph
            client_id = self.client_id  # Replace with your app's client ID
            tenant_id = self.tenant_id  # Replace with your tenant ID
            authority = f"https://login.microsoftonline.com/{tenant_id}"
            scopes = ["User.Read", "GroupMember.Read.All"]  # Combine necessary scopes

            # Initialize the MSAL application
            app = PublicClientApplication(client_id, authority=authority)

            # Check the cache for an existing token
            accounts = app.get_accounts()
            if accounts:
                token_result = app.acquire_token_silent(scopes, account=accounts[0])
            else:
                # Prompt for login if no token is available
                token_result = app.acquire_token_interactive(scopes=scopes)

            # Retrieve the access token and user details
            access_token = token_result.get("access_token")
            if not access_token:
                raise Exception("Failed to acquire access token.")

            upn = token_result["id_token_claims"]["preferred_username"]
            print(f"Authenticated user: {upn}")

            # Query Microsoft Graph for group memberships
            headers = {"Authorization": f"Bearer {access_token}"}
            graph_url = f"https://graph.microsoft.com/v1.0/users/{upn}/memberOf"
            response = requests.get(graph_url, headers=headers)
            response.raise_for_status()

            # Check if the user belongs to any required groups
            required_groups = ["InventorySystemUser", "InventorySystemAdmin", "InventorySystemSuperAdmin"]
            groups = response.json().get("value", [])
            for group in groups:
                if group.get("displayName") in required_groups:
                    role = group.get("displayName")
                    print(f"User {upn} is authorized with role: {role}")
                    return upn, role  # Return UPN and role if authorized

            print(f"User {upn} is not authorized.")
            return None, None  # User not in required groups

        except Exception as e:
            print(f"Error during authentication or access check: {e}")
            return None, None """
        
    def selection_window(self):
        self.selection_frame = tk.Frame(self.root, padx=20, pady=20, bg="#3E4C59")
        self.selection_frame.pack(expand=True, fill='both', padx=10, pady=10)

        self.selection_label = tk.Label(
            self.selection_frame,
            text="Select Inventory Type",
            font=("Arial", 14, "bold"),
            fg="#E4E7EB",
            bg="#3E4C59"
        )
        self.selection_label.pack(pady=10)

        # Hardware Inventory Button
        self.hardware_button = tk.Button(
            self.selection_frame,
            text="Hardware Inventory",
            command=self.start_hardware_inventory,
            font=("Arial", 12),
            bg="#4A5568",
            fg="#E4E7EB",
            activebackground="#6B7280"
        )
        self.hardware_button.pack(pady=10, fill='x')

        # Software Inventory Button
        self.software_button = tk.Button(
            self.selection_frame,
            text="Software Inventory",
            command=self.start_software_inventory,  # Correctly linked
            font=("Arial", 12),
            bg="#4A5568",
            fg="#E4E7EB",
            activebackground="#6B7280"
        )
        self.software_button.pack(pady=10, fill='x')

    """ def selection_window(self):
        upn, role = self.authenticate_and_check_access()
        if not upn or not role:
            # Show error screen if the user is unauthorized
            self.show_error_screen(
                "Unauthorized Access",
                "You do not have the required permissions to access the Inventory System."
            )
            return

        # Create the selection frame
        self.selection_frame = tk.Frame(self.root, padx=20, pady=20, bg="#3E4C59")
        self.selection_frame.pack(expand=True, fill='both', padx=10, pady=10)

        # Display the signed-in user's UPN and role
        user_label = tk.Label(
            self.selection_frame,
            text=f"User: {upn}",
            font=("Arial", 10, "bold"),
            fg="#E4E7EB",
            bg="#3E4C59"
        )
        user_label.pack(pady=5)

        role_label = tk.Label(
            self.selection_frame,
            text=f"Role: {role}",
            font=("Arial", 10, "bold"),
            fg="#E4E7EB",
            bg="#3E4C59"
        )
        role_label.pack(pady=5)

        # Selection label
        self.selection_label = tk.Label(
            self.selection_frame,
            text="Select Inventory Type",
            font=("Arial", 14, "bold"),
            fg="#E4E7EB",
            bg="#3E4C59"
        )
        self.selection_label.pack(pady=10)

        # Hardware inventory button
        self.hardware_button = tk.Button(
            self.selection_frame,
            text="Hardware Inventory",
            command=self.start_hardware_inventory,
            font=("Arial", 12),
            bg="#4A5568",
            fg="#E4E7EB",
            activebackground="#6B7280"
        )
        self.hardware_button.pack(pady=10, fill='x')

        # Software inventory button
        self.software_button = tk.Button(
            self.selection_frame,
            text="Software Inventory",
            command=self.start_software_inventory,
            font=("Arial", 12),
            bg="#4A5568",
            fg="#E4E7EB",
            activebackground="#6B7280"
        )
        self.software_button.pack(pady=10, fill='x') """

    """ def show_error_screen(self, title, message):
        error_frame = tk.Frame(self.root, padx=20, pady=20, bg="#3E4C59")
        error_frame.pack(expand=True, fill='both', padx=10, pady=10)

        error_label = tk.Label(error_frame, text=title, font=("Arial", 16, "bold"), fg="#FF0000", bg="#3E4C59")
        error_label.pack(pady=10)

        message_label = tk.Label(error_frame, text=message, font=("Arial", 12), fg="#E4E7EB", bg="#3E4C59", wraplength=400, justify="center")
        message_label.pack(pady=10)

        quit_button = tk.Button(error_frame, text="Quit", command=self.root.quit, font=("Arial", 12), bg="#4A5568", fg="#E4E7EB", activebackground="#6B7280")
        quit_button.pack(pady=20) """

    def load_software_data(self):
        """
        Load software data from Intune API, store it in the database, and display it in the TreeView.
        """
        try:
            # Step 1: Authenticate and retrieve data from Intune
            access_token = self.get_access_token()  # Ensure this method exists to get the token
            headers = {"Authorization": f"Bearer {access_token}"}
            url = "https://graph.microsoft.com/v1.0/deviceManagement/detectedApps/"
            
            all_software_data = []
            while url:
                response = self.session.get(url, headers=headers, verify=False, proxies=self.PROXIES)
                if response.status_code == 200:
                    data = response.json()
                    all_software_data.extend(data.get("value", []))
                    url = data.get("@odata.nextLink")  # Check for the next page of data
                else:
                    raise Exception(f"Failed to fetch software data: {response.status_code}, {response.text}")

            # Step 2: Process and write data into the database
            with sqlite3.connect("inventory.db", timeout=30) as conn:
                cursor = conn.cursor()

                # Ensure the Software table exists
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS Software (
                        SoftwareName TEXT,
                        Version TEXT,
                        InstalledDevices INTEGER,
                        PRIMARY KEY (SoftwareName, Version)
                    )
                ''')

                for app in all_software_data:
                    software_name = app.get("displayName", "").strip()
                    version = app.get("version", "").strip()
                    installed_devices = app.get("deviceCount", 0)

                    if software_name and version:
                        cursor.execute('''
                            INSERT INTO Software (SoftwareName, Version, InstalledDevices)
                            VALUES (?, ?, ?)
                            ON CONFLICT(SoftwareName, Version)
                            DO UPDATE SET InstalledDevices = excluded.InstalledDevices
                        ''', (software_name, version, installed_devices))

                conn.commit()

            # Step 3: Load data from the database
            with sqlite3.connect("inventory.db", timeout=30) as conn:
                software_data = pd.read_sql_query("SELECT * FROM Software", conn)

            # Step 4: Store and display the data
            self.software_data = software_data
            self.filtered_software_data = software_data
            self.current_software_page = 0

            # Display the first page of data
            self.display_software_data()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load software data: {e}")

    def next_software_page(self):
        total_pages = (len(self.filtered_software_data) - 1) // self.software_page_size + 1
        if self.current_software_page + 1 < total_pages:
            self.current_software_page += 1
            self.display_software_data()

    def previous_software_page(self):
        if self.current_software_page > 0:
            self.current_software_page -= 1
            self.display_software_data()

    def apply_software_filter(self):
        try:
            # Get the search query
            search_query = self.software_search_entry.get().strip().lower()

            # Start with the full dataset
            filtered_data = self.software_data.copy()

            # Apply search query
            if search_query:
                filtered_data = filtered_data[
                    filtered_data.apply(
                        lambda row: search_query in " ".join(row.astype(str).str.lower()), axis=1
                    )
                ]

            # Update the filtered data and display it
            self.filtered_software_data = filtered_data
            self.current_software_page = 0
            self.display_software_data()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to apply filter: {e}")

    def clear_software_filter(self):
        try:
            # Clear search entry
            self.software_search_entry.delete(0, tk.END)

            # Reset the filtered data and display it
            self.filtered_software_data = self.software_data
            self.current_software_page = 0
            self.display_software_data()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to clear filters: {e}")

    def export_software_view(self):
        """
        Export the filtered software view to an Excel file.
        """
        try:
            # Prompt the user to select a save location
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Export Software View",
            )
            if file_path:
                # Export the filtered data to an Excel file
                self.filtered_software_data.to_excel(file_path, index=False)
                messagebox.showinfo("Success", "Software view exported successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export software view: {e}")

    def export_software_database(self):
        """
        Export the entire software database to an Excel file.
        """
        try:
            # Prompt the user to select a save location
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Export Software Database",
            )
            if file_path:
                # Fetch the entire Software table from the database
                with sqlite3.connect('inventory.db', timeout=30) as conn:
                    software_data = pd.read_sql_query("SELECT * FROM Software", conn)

                # Export the data to an Excel file
                software_data.to_excel(file_path, index=False)
                messagebox.showinfo("Success", "Software database exported successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export software database: {e}")

    def display_software_data(self):
        """
        Display the software data in the TreeView for the current page.
        """
        try:
            # Clear the existing rows in the TreeView
            for item in self.software_tree.get_children():
                self.software_tree.delete(item)

            # Determine the range of rows to display for the current page
            start_idx = self.current_software_page * self.software_page_size
            end_idx = start_idx + self.software_page_size
            current_page_data = self.filtered_software_data.iloc[start_idx:end_idx]

            # Populate the TreeView with the current page of data
            for _, row in current_page_data.iterrows():
                self.software_tree.insert("", "end", values=(row["SoftwareName"], row["Version"], row["InstalledDevices"]))

            # Update the pagination label
            total_pages = (len(self.filtered_software_data) - 1) // self.software_page_size + 1
            self.softwarepage_number_label.config(text=f"Page {self.current_software_page + 1} / {total_pages}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to display software data: {e}")
    
    def clear_software_search(self):
        """
        Clear the search entry and reload the full software data.
        """
        self.software_search_entry.delete(0, tk.END)  # Clear the search entry field
        self.filtered_software_data = self.software_data.copy()  # Reset filtered data to full data
        self.current_software_page = 0  # Reset pagination
        self.display_software_data()  # Refresh the Treeview with full data

    def refresh_software_data(self):
        """
        Refresh the software data by reloading it from the database and API.
        """
        try:
            # Refresh software data from the database or API
            self.load_software_data()

            # Reset filters and pagination
            self.clear_software_filter()
            self.current_software_page = 0

            # Display the updated software data
            self.display_software_data()

            messagebox.showinfo("Success", "Software data refreshed successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to refresh software data: {e}")
    
    def start_software_inventory(self):
        # Create a new window for the progress bar
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Starting Software Inventory")
        progress_window.geometry("300x100")
        progress_window.transient(self.root)
        progress_window.grab_set()  # Make the main window unresponsive

        # Center the progress window on the screen
        progress_window.update_idletasks()
        x = (progress_window.winfo_screenwidth() // 2) - (progress_window.winfo_width() // 2)
        y = (progress_window.winfo_screenheight() // 2) - (progress_window.winfo_height() // 2)
        progress_window.geometry(f"+{x}+{y}")

        # Add a progress bar widget
        progress = ttk.Progressbar(progress_window, orient="horizontal", length=250, mode="indeterminate")
        progress.pack(pady=20)
        progress.start()

        # Run the hardware inventory in a separate thread to keep the UI responsive
        threading.Thread(target=self.run_software_inventory, args=(progress_window, progress)).start()

    def run_software_inventory(self, progress_window, progress):
        try:
            # Step 1: Authenticate to Microsoft Graph to retrieve Intune data
            access_token = self.get_access_token()
            
            # Step 2: Retrieve Intune apps information
            intune_apps = self.retrieve_intune_software(access_token)

            # Step 3: Export the retrieved app information to the database
            self.export_apps_to_database(intune_apps)
        except Exception as e:
            print(f"Failed to run software inventory: {e}\n{traceback.format_exc()}")
        finally:
            # Stop and close the progress window
            progress.stop()
            progress_window.destroy()

        # Destroy the selection frame and load the main app
        self.selection_frame.destroy()
        self.main_app("software")

    def retrieve_intune_software(self, access_token):
        url = "https://graph.microsoft.com/v1.0/deviceManagement/detectedApps/"
        
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        response = requests.get(url, headers=headers, verify=False, proxies=self.PROXIES)
        if response.status_code == 200:
            apps = response.json().get('value', [])            
            return apps
        else:
            raise Exception(f"Failed to retrieve software information. Status code: {response.status_code}, Response: {response.text}")
        
    def export_apps_to_database(self, apps):
        # Prepare a list of selected app data for saving to Excel
        selected_apps = []
        for app in apps:
            app_data = {
                'SoftwareName': app.get('displayName', ''),
                'Version': app.get('version', ''),
                'InstalledDevices': app.get('deviceCount', '')
            }

            # Add a check to skip empty or nearly-empty records
            if not any(app_data.values()):
                print("Skipping empty app record:", app_data)
                continue
            
            selected_apps.append(app_data)

    def start_hardware_inventory(self):
        # Create a new window for the progress bar
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Starting Hardware Inventory")
        progress_window.geometry("300x100")
        progress_window.transient(self.root)
        progress_window.grab_set()  # Make the main window unresponsive

        # Center the progress window on the screen
        progress_window.update_idletasks()
        x = (progress_window.winfo_screenwidth() // 2) - (progress_window.winfo_width() // 2)
        y = (progress_window.winfo_screenheight() // 2) - (progress_window.winfo_height() // 2)
        progress_window.geometry(f"+{x}+{y}")

        # Add a progress bar widget
        progress = ttk.Progressbar(progress_window, orient="horizontal", length=250, mode="indeterminate")
        progress.pack(pady=20)
        progress.start()

        # Run the hardware inventory in a separate thread to keep the UI responsive
        threading.Thread(target=self.run_hardware_inventory, args=(progress_window, progress)).start()

    def run_hardware_inventory(self, progress_window, progress):
        try:
            #Export the retrieved device information to the database
            self.export_to_database()
        except Exception as e:
            print(f"Failed to run hardware inventory: {e}\n{traceback.format_exc()}")
        finally:
            # Stop and close the progress window
            progress.stop()
            progress_window.destroy()

        # Destroy the selection frame and load the main app
        self.selection_frame.destroy()
        self.main_app("hardware")

    def return_to_main_screen(self):
        # Destroy the main app frame and return to the selection window
        for widget in self.root.winfo_children():
            widget.destroy()
        self.set_window_size()
        self.selection_window()

    def clear_search_entry(self):
        self.search_entry.delete(0, tk.END)
        # If there are active filters, apply them, otherwise show the full dataset
        if any(self.filter_column_comboboxes[i].get() for i in range(len(self.filter_column_comboboxes))):
            self.apply_filters()
        else:
            self.filtered_data = self.data.copy()
            self.current_page = 0  # Reset to the first page
            self.display_data(self.filtered_data)
            self.update_total_records(self.filtered_data)
            self.update_page_number_label()  # Update the page number label

    def execute_query(self, query, params=()):
        retries = 10
        delay = 1
        while retries > 0:
            with self.db_lock:
                try:
                    self.cursor.execute(query, params)
                    break
                except sqlite3.OperationalError as e:
                    if 'database is locked' in str(e):
                        time.sleep(delay)
                        delay *= 2  # Exponential backoff
                        retries -= 1
                    else:
                        raise
        if retries == 0:
            raise sqlite3.OperationalError("Database is locked after multiple retries.")
        if threading.get_ident() == self.conn_thread_id:
            self.cursor.execute(query, params)
        else:
            with sqlite3.connect('inventory.db', timeout=30) as conn:
                cursor = conn.cursor()
                cursor.execute(query, params)

    def fetchall_query(self, query, params=()):
        retries = 10
        delay = 1
        while retries > 0:
            with self.db_lock:
                try:
                    self.cursor.execute(query, params)
                    return self.cursor.fetchall()
                except sqlite3.OperationalError as e:
                    if 'database is locked' in str(e):
                        time.sleep(delay)
                        delay *= 2  # Exponential backoff
                        retries -= 1
                    else:
                        raise
        if retries == 0:
            raise sqlite3.OperationalError("Database is locked after multiple retries.")
        if threading.get_ident() == self.conn_thread_id:
            return self.fetchall_query("SELECT * FROM Remarks")
        else:
            with sqlite3.connect('inventory.db') as conn:
                cursor = conn.cursor()
                return cursor.fetchall()

        if retries == 0:
            raise sqlite3.OperationalError("Database is locked after multiple retries.")
        if threading.get_ident() == self.conn_thread_id:
            self.conn.commit()
        else:
            with sqlite3.connect('inventory.db') as conn:
                conn.commit()

    def check_thread_and_execute(self, query, params=()):
        retries = 5
        while retries > 0:
            with self.db_lock:
                try:
                    self.cursor.execute(query, params)
                    break
                except sqlite3.OperationalError as e:
                    if 'database is locked' in str(e):
                        time.sleep(1)
                        retries -= 1
                    else:
                        raise
                else:
                    raise
                    
    def load_data(self):
        # Create a new window for the progress bar
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Loading Data")
        progress_window.geometry("300x100")
        progress_window.transient(self.root)
        progress_window.grab_set()  # Make the main window unresponsive

        # Center the progress window on the screen
        progress_window.update_idletasks()
        x = (progress_window.winfo_screenwidth() // 2) - (progress_window.winfo_width() // 2)
        y = (progress_window.winfo_screenheight() // 2) - (progress_window.winfo_height() // 2)
        progress_window.geometry(f"+{x}+{y}")

        # Add a progress bar widget
        progress = ttk.Progressbar(progress_window, orient="horizontal", length=250, mode="indeterminate")
        progress.pack(pady=20)
        progress.start()

        file_path = r"DeviceInventory.xlsx"
        try:
            # Step 1: Load data from the Devices table in the database
            with sqlite3.connect('inventory.db', timeout=30) as conn:
                devices_data = pd.read_sql_query("SELECT * FROM Devices", conn)
                devices_data = devices_data.fillna("")

            # Step 2: Load data from Excel file
            excel_data = pd.read_excel(file_path)
            excel_data.dropna(subset=['DeviceName', 'SerialNumber'], how='all', inplace=True)
            excel_data = excel_data.fillna("")

            # Convert SerialNumber and DeviceName to string to handle numeric values
            excel_data['SerialNumber'] = excel_data['SerialNumber'].astype(str)
            excel_data['DeviceName'] = excel_data['DeviceName'].astype(str)
            devices_data['SerialNumber'] = devices_data['SerialNumber'].astype(str)
            devices_data['DeviceName'] = devices_data['DeviceName'].astype(str)

            # Ensure Remarks column exists in both DataFrames
            if 'Remarks' not in excel_data.columns:
                excel_data['Remarks'] = ""

            if 'Remarks' not in devices_data.columns:
                devices_data['Remarks'] = ""

            # Combine data from Devices table and Excel data
            if not devices_data.empty and not excel_data.empty:
                combined_data = pd.concat([devices_data, excel_data], ignore_index=True)
            else:
                combined_data = devices_data if not devices_data.empty else excel_data

            # Remove duplicates based on DeviceName and SerialNumber, keeping the latest ReportTime
            combined_data['ReportTime'] = pd.to_datetime(combined_data['ReportTime'], format="%Y-%m-%d %H:%M:%S", errors='coerce')
            combined_data.sort_values(by=['DeviceName', 'SerialNumber', 'ReportTime'], ascending=[True, True, False], inplace=True)
            combined_data.dropna(subset=['DeviceName', 'SerialNumber'], inplace=True)  # Drop rows with missing DeviceName or SerialNumber
            combined_data.drop_duplicates(subset=['DeviceName', 'SerialNumber'], keep='first', inplace=True)

            # Fill NaN values with empty strings
            combined_data.fillna("", inplace=True)

            # Load Remarks from the Remarks table and update the DataFrame
            with sqlite3.connect('inventory.db', timeout=30) as conn:
                remarks_df = pd.read_sql_query("SELECT * FROM Remarks", conn)
                remarks_df['SerialNumber'] = remarks_df['SerialNumber'].astype(str).str.strip().str.lower()
                remarks_df['DeviceName'] = remarks_df['DeviceName'].astype(str).str.strip().str.lower()
                remarks_df.set_index(['SerialNumber', 'DeviceName'], inplace=True)

                def merge_remarks(row):
                    key = (str(row['SerialNumber']).strip().lower(), str(row['DeviceName']).strip().lower())
                    if key in remarks_df.index:
                        return remarks_df.at[key, 'Remarks'] if remarks_df.at[key, 'Remarks'] else row['Remarks']
                    else:
                        return row['Remarks']

                combined_data['Remarks'] = combined_data.apply(merge_remarks, axis=1)

            # Update Source column based on Remarks
            if 'Source' not in combined_data.columns:
                combined_data['Source'] = 'Cloud'

            # Identify records in the database but not in the latest Excel data and update Source to "Local"
            latest_records = set(zip(excel_data['SerialNumber'].str.strip().str.lower(), excel_data['DeviceName'].str.strip().str.lower()))
            combined_data['SerialNumber_lower'] = combined_data['SerialNumber'].str.strip().str.lower()
            combined_data['DeviceName_lower'] = combined_data['DeviceName'].str.strip().str.lower()

            # Set Source to 'Intune' for records present in the latest data
            combined_data.loc[combined_data.apply(lambda row: (row['SerialNumber_lower'], row['DeviceName_lower']) in latest_records, axis=1), 'Source'] = 'Cloud'

            # Set Source to 'Local' for records that are not in the latest data
            combined_data.loc[~combined_data.apply(lambda row: (row['SerialNumber_lower'], row['DeviceName_lower']) in latest_records, axis=1), 'Source'] = 'Local'

            # Assign combined_data to self.data and self.filtered_data, then display in Treeview
            self.data = combined_data.drop(columns=['SerialNumber_lower', 'DeviceName_lower'])
            self.filtered_data = self.data.copy()
            self.current_page = 0  # Ensure current page starts from 0
            self.display_data(self.filtered_data)
            self.update_total_records(self.filtered_data)

            # Write the loaded data into the Devices table in the database
            with sqlite3.connect('inventory.db', timeout=30) as conn:
                cursor = conn.cursor()

                for _, row in self.data.iterrows():
                    # Convert ReportTime to string or None (if it's NaT)
                    report_time_str = row['ReportTime']
                    if pd.isna(report_time_str):
                        report_time_str = None
                    elif isinstance(report_time_str, pd.Timestamp):
                        report_time_str = report_time_str.strftime("%Y-%m-%d %H:%M:%S")

                    # Check if the SerialNumber already exists in the database
                    cursor.execute(
                        "SELECT * FROM Devices WHERE DeviceName = ? AND SerialNumber = ?", 
                        (row['DeviceName'], row['SerialNumber'])
                    )
                    existing_record = cursor.fetchone()

                    if existing_record:
                        # Update the existing record and set Source to "Cloud" if the record is found in the latest data
                        source_value = 'Cloud' if (str(row['SerialNumber']).strip().lower(), str(row['DeviceName']).strip().lower()) in latest_records else 'Local'

                        # Update the existing record
                        cursor.execute('''UPDATE Devices SET 
                                        UserPrincipalName = ?, OperatingSystem = ?, EntraDeviceID = ?, IntuneDeviceID = ?, OSVersion = ?, ComplianceState = ?, Source = ?,
                                        Model = ?, Manufacturer = ?, MAC = ?, IntuneLastSync = ?, ReportTime = ?, Encryption = ?,
                                        UserDisplayName = ?, JobTitle = ?, Department = ?, City = ?, Country = ?,
                                        TrustType = ?, TotalStorage = ?, FreeStorage = ?, PhysicalMemory = ?
                                        WHERE DeviceName = ? AND SerialNumber = ?''',
                                    (row['UserPrincipalName'], row['OperatingSystem'], row['EntraDeviceID'], row['IntuneDeviceID'], row['OSVersion'], row['ComplianceState'], source_value,
                                    row['Model'], row['Manufacturer'], row['MAC'], row['IntuneLastSync'], report_time_str, row['Encryption'],
                                    row['UserDisplayName'], row['JobTitle'], row['Department'], row['City'], row['Country'],
                                    row['TrustType'], row['TotalStorage'], row['FreeStorage'], row['PhysicalMemory'],
                                    row['DeviceName'], row['SerialNumber']))
                    else:
                        # Insert a new record
                        cursor.execute('''INSERT INTO Devices (DeviceName, SerialNumber, UserPrincipalName, EntraDeviceID, IntuneDeviceID, OperatingSystem, OSVersion, Source, ComplianceState, 
                                       Model, Manufacturer, MAC, IntuneLastSync, ReportTime, Encryption, UserDisplayName, JobTitle, Department, City, Country,
                                       TrustType, TotalStorage, FreeStorage, PhysicalMemory)
                                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                                    (row['DeviceName'], row['SerialNumber'], row['UserPrincipalName'], row['EntraDeviceID'], row['IntuneDeviceID'], 
                                    row['OperatingSystem'], row['OSVersion'], 'Cloud', row['ComplianceState'], row['Model'], 
                                    row['Manufacturer'], row['MAC'], row['IntuneLastSync'], report_time_str, row['Encryption'], row['UserDisplayName'], 
                                    row['JobTitle'], row['Department'], row['City'], row['Country'],
                                    row['TrustType'], row['TotalStorage'], row['FreeStorage'], row['PhysicalMemory']))

                conn.commit()

            # Update the page number label after loading data
            self.update_page_number_label()

        except Exception as e:
            print(f"Failed to load data: {e}\n{traceback.format_exc()}")
        finally:
            # Stop the progress bar and close the progress window
            progress.stop()
            progress_window.destroy()

    def load_data_for_refresh(self):        
        file_path = r"DeviceInventory.xlsx"
        try:
            # Step 1: Load data from the Devices table in the database
            with sqlite3.connect('inventory.db', timeout=30) as conn:
                devices_data = pd.read_sql_query("SELECT * FROM Devices", conn)
                devices_data = devices_data.fillna("")

            # Step 2: Load data from Excel file
            excel_data = pd.read_excel(file_path)
            excel_data.dropna(subset=['DeviceName', 'SerialNumber'], how='all', inplace=True)
            excel_data = excel_data.fillna("")

            # Convert SerialNumber and DeviceName to string to handle numeric values
            excel_data['SerialNumber'] = excel_data['SerialNumber'].astype(str)
            excel_data['DeviceName'] = excel_data['DeviceName'].astype(str)
            devices_data['SerialNumber'] = devices_data['SerialNumber'].astype(str)
            devices_data['DeviceName'] = devices_data['DeviceName'].astype(str)

            # Ensure Remarks column exists in both DataFrames
            if 'Remarks' not in excel_data.columns:
                excel_data['Remarks'] = ""

            if 'Remarks' not in devices_data.columns:
                devices_data['Remarks'] = ""

            # Combine data from Devices table and Excel data
            if not devices_data.empty and not excel_data.empty:
                combined_data = pd.concat([devices_data, excel_data], ignore_index=True)
            else:
                combined_data = devices_data if not devices_data.empty else excel_data

            # Remove duplicates based on DeviceName and SerialNumber, keeping the latest ReportTime
            combined_data['ReportTime'] = pd.to_datetime(combined_data['ReportTime'], format="%Y-%m-%d %H:%M:%S", errors='coerce')
            combined_data.sort_values(by=['DeviceName', 'SerialNumber', 'ReportTime'], ascending=[True, True, False], inplace=True)
            combined_data.dropna(subset=['DeviceName', 'SerialNumber'], inplace=True)  # Drop rows with missing DeviceName or SerialNumber
            combined_data.drop_duplicates(subset=['DeviceName', 'SerialNumber'], keep='first', inplace=True)

            # Fill NaN values with empty strings
            combined_data.fillna("", inplace=True)

            # Load Remarks from the Remarks table and update the DataFrame
            with sqlite3.connect('inventory.db', timeout=30) as conn:
                remarks_df = pd.read_sql_query("SELECT * FROM Remarks", conn)
                remarks_df['SerialNumber'] = remarks_df['SerialNumber'].astype(str).str.strip().str.lower()
                remarks_df['DeviceName'] = remarks_df['DeviceName'].astype(str).str.strip().str.lower()
                remarks_df.set_index(['SerialNumber', 'DeviceName'], inplace=True)

                def merge_remarks(row):
                    key = (str(row['SerialNumber']).strip().lower(), str(row['DeviceName']).strip().lower())
                    if key in remarks_df.index:
                        return remarks_df.at[key, 'Remarks'] if remarks_df.at[key, 'Remarks'] else row['Remarks']
                    else:
                        return row['Remarks']

                combined_data['Remarks'] = combined_data.apply(merge_remarks, axis=1)

            # Update Source column based on Remarks
            if 'Source' not in combined_data.columns:
                combined_data['Source'] = 'Cloud'

            # Identify records in the database but not in the latest Excel data and update Source to "Local"
            latest_records = set(zip(excel_data['SerialNumber'].str.strip().str.lower(), excel_data['DeviceName'].str.strip().str.lower()))
            combined_data['SerialNumber_lower'] = combined_data['SerialNumber'].str.strip().str.lower()
            combined_data['DeviceName_lower'] = combined_data['DeviceName'].str.strip().str.lower()

            # Set Source to 'Cloud' for records present in the latest data
            combined_data.loc[combined_data.apply(lambda row: (row['SerialNumber_lower'], row['DeviceName_lower']) in latest_records, axis=1), 'Source'] = 'Cloud'

            # Set Source to 'Local' for records that are not in the latest data
            combined_data.loc[~combined_data.apply(lambda row: (row['SerialNumber_lower'], row['DeviceName_lower']) in latest_records, axis=1), 'Source'] = 'Local'

            # Assign combined_data to self.data and self.filtered_data, then display in Treeview
            self.data = combined_data.drop(columns=['SerialNumber_lower', 'DeviceName_lower'])
            self.filtered_data = self.data.copy()
            self.current_page = 0  # Ensure current page starts from 0
            self.display_data(self.filtered_data)
            self.update_total_records(self.filtered_data)

            # Write the loaded data into the Devices table in the database
            with sqlite3.connect('inventory.db', timeout=30) as conn:
                cursor = conn.cursor()

                for _, row in self.data.iterrows():
                    # Convert ReportTime to string or None (if it's NaT)
                    report_time_str = row['ReportTime']
                    if pd.isna(report_time_str):
                        report_time_str = None
                    elif isinstance(report_time_str, pd.Timestamp):
                        report_time_str = report_time_str.strftime("%Y-%m-%d %H:%M:%S")

                    # Check if the SerialNumber already exists in the database
                    cursor.execute(
                        "SELECT * FROM Devices WHERE DeviceName = ? AND SerialNumber = ?", 
                        (row['DeviceName'], row['SerialNumber'])
                    )
                    existing_record = cursor.fetchone()

                    if existing_record:
                        # Update the existing record and set Source to "Cloud" if the record is found in the latest data
                        source_value = 'Cloud' if (str(row['SerialNumber']).strip().lower(), str(row['DeviceName']).strip().lower()) in latest_records else 'Local'

                        # Update the existing record
                        cursor.execute('''UPDATE Devices SET 
                                        UserPrincipalName = ?, OperatingSystem = ?, EntraDeviceID = ?, IntuneDeviceID = ?, OSVersion = ?, ComplianceState = ?, Source = ?,
                                        Model = ?, Manufacturer = ?, MAC = ?, IntuneLastSync = ?, ReportTime = ?, Encryption = ?,
                                        UserDisplayName = ?, JobTitle = ?, Department = ?, City = ?, Country = ?,
                                        TrustType = ?, TotalStorage = ?, FreeStorage = ?, PhysicalMemory = ?
                                        WHERE DeviceName = ? AND SerialNumber = ?''',
                                    (row['UserPrincipalName'], row['OperatingSystem'], row['EntraDeviceID'], row['IntuneDeviceID'], row['OSVersion'], row['ComplianceState'], source_value,
                                    row['Model'], row['Manufacturer'], row['MAC'], row['IntuneLastSync'], report_time_str, row['Encryption'],
                                    row['UserDisplayName'], row['JobTitle'], row['Department'], row['City'], row['Country'],
                                    row['TrustType'], row['TotalStorage'], row['FreeStorage'], row['PhysicalMemory'],
                                    row['DeviceName'], row['SerialNumber']))
                    else:
                        # Insert a new record
                        cursor.execute('''INSERT INTO Devices (DeviceName, SerialNumber, UserPrincipalName, EntraDeviceID, IntuneDeviceID, OperatingSystem, OSVersion, Source, ComplianceState, 
                                       Model, Manufacturer, MAC, IntuneLastSync, ReportTime, Encryption, UserDisplayName, JobTitle, Department, City, Country,
                                       TrustType, TotalStorage, FreeStorage, PhysicalMemory)
                                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                                    (row['DeviceName'], row['SerialNumber'], row['UserPrincipalName'], row['EntraDeviceID'], row['IntuneDeviceID'], 
                                    row['OperatingSystem'], row['OSVersion'], 'Cloud', row['ComplianceState'], row['Model'], 
                                    row['Manufacturer'], row['MAC'], row['IntuneLastSync'], report_time_str, row['Encryption'], row['UserDisplayName'], 
                                    row['JobTitle'], row['Department'], row['City'], row['Country'],
                                    row['TrustType'], row['TotalStorage'], row['FreeStorage'], row['PhysicalMemory']))

                conn.commit()

            # Update the page number label after loading data
            self.update_page_number_label()

        except Exception as e:
            print(f"Failed to load data: {e}\n{traceback.format_exc()}")

    def get_access_token(self):
        url = f'https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token'
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }
        body = {
            'grant_type': 'client_credentials',
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'scope': 'https://graph.microsoft.com/.default'
        }
        try:
            response = self.session.post(url, headers=headers, data=body, verify=False, proxies=self.PROXIES)
            response.raise_for_status()
            return response.json().get('access_token')
        except requests.exceptions.ProxyError:
            raise Exception("Proxy error occurred. Please check your proxy settings.")
        except requests.exceptions.RequestException as e:
            raise Exception(f"Failed to acquire access token: {e}")

    def retrieve_intune_devices(self, access_token):
        url = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
        
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        response = requests.get(url, headers=headers, verify=False, proxies=self.PROXIES)
        if response.status_code == 200:
            devices = response.json().get('value', [])
            # Retrieve user details for each device
            for device in devices:
                user_principal_name = device.get('userPrincipalName', '')
                device_id = device.get('azureADDeviceId', '')
                intune_device_id = device.get('id', '')
                if user_principal_name:
                    self.retrieve_user_details(access_token, user_principal_name, device)
                if device_id:
                    self.retrieve_entra_devices(access_token, device_id, device)
                    #self.retrieve_device_memory(access_token, intune_device_id, device)
            return devices
        else:
            raise Exception(f"Failed to retrieve device information. Status code: {response.status_code}, Response: {response.text}")
        
    def retrieve_entra_devices(self, access_token, deviceId, device):
        # Get user details
        entra_info_url = f"https://graph.microsoft.com/v1.0/devices?$filter=deviceId eq '{deviceId}'&$select=createdDateTime,trustType"
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }

        try:
            # Request entra device details
            entra_info_response = requests.get(entra_info_url, headers=headers, verify=False, proxies=self.PROXIES)
            if entra_info_response.status_code == 200:
                device_list = entra_info_response.json().get('value', [])  # Extract 'value' as a list
                if device_list:  # Ensure the list is not empty
                    first_device = device_list[0]  # Extract the first device from the list
                    trust_type = first_device.get('trustType', 'N/A')

                    # Update device dictionary
                    device['trustType'] = trust_type
                else:
                    print(f"No Entra device found for DeviceID: {deviceId}")
            else:
                print(f"Failed to retrieve Entra device details. Status code: {entra_info_response.status_code}, Response: {entra_info_response.text}")
        except Exception as e:
            print(f"Exception occurred while retrieving Entra device details: {str(e)}")

    # To be implemented
    def refresh_device_info(self):
        selected_item = self.tree.selection()
        if selected_item:
            item_values = self.tree.item(selected_item, "values")
            # Create a mapping of column names to their index positions
            columns_mapping = {col: idx for idx, col in enumerate(self.tree["columns"])}
            device_name = str(item_values[columns_mapping["DeviceName"]])  
            serial_number = str(item_values[columns_mapping["SerialNumber"]])  

            # Retrieve DeviceID from the database
            self.cursor.execute("SELECT EntraDeviceID FROM Devices WHERE DeviceName = ? AND SerialNumber = ?", (device_name, serial_number))
            result = self.cursor.fetchone()
            if result:
                entra_device_id = result[0]
            else:
                messagebox.showerror("Error", "Entra Device ID not found for the selected device.")

            self.cursor.execute("SELECT IntuneDeviceID FROM Devices WHERE DeviceName = ? AND SerialNumber = ?", (device_name, serial_number))
            result = self.cursor.fetchone()
            if result:
                intune_device_id = result[0]
            else:
                messagebox.showerror("Error", "Intune Device ID not found for the selected device.")    

            try:
                # Authenticate and get the access token
                access_token = self.get_access_token()
                headers = {
                    'Authorization': f'Bearer {access_token}',
                    'Content-Type': 'application/json'
                }

                # Retrieve device details from Microsoft Graph API
                device_url = f"https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/{intune_device_id}"
                response = requests.get(device_url, headers=headers, verify=False, proxies=self.PROXIES)
                print("Intune Device Info Retrieved")
                if response.status_code == 200:
                    device_data = response.json()
                    total_storage = device_data.get('totalStorageSpaceInBytes', '')
                    free_storage = device_data.get('freeStorageSpaceInBytes', '')

                    # Convert to GB and round to the nearest integer if the value is not 'N/A'
                    total_storage_gb = str(int(round(total_storage / (1024 ** 3))) if isinstance(total_storage, (int, float)) and total_storage != '' else '')
                    free_storage_gb = str(int(round(free_storage / (1024 ** 3))) if isinstance(free_storage, (int, float)) and free_storage != '' else '')
                    print("Total Storage: " + total_storage_gb)
                    print("Total Storage: " + free_storage_gb)
                    updated_values = {
                        "OperatingSystem": device_data.get("operatingSystem", ""),
                        "OSVersion": device_data.get("osVersion", ""),
                        "ComplianceState": device_data.get("complianceState", ""),
                        "Manufacturer": device_data.get("manufacturer", ""),
                        "Model": device_data.get("model", ""),
                        "IntuneLastSync": device_data.get("lastSyncDateTime", ""),                
                        'EntraDeviceID': device_data.get('EntraDeviceID', ""),
                        'IntuneDeviceID': device_data.get('IntuneDeviceID', ""),
                        'UserPrincipalName': device_data.get('UserPrincipalName', ""),                        
                        'MAC': device_data.get('MAC', ""),                        
                        'ReportTime': time.strftime("%Y-%m-%d %H:%M:%S"),
                        'Source': "Cloud",
                        'Encryption': device_data.get('Encryption', ""),                        
                        'TotalStorage': total_storage_gb,
                        'FreeStorage': free_storage_gb
                    }

                    # Retrieve physical memory details from the separate API
                    memory_url = f"https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/{intune_device_id}?$select=physicalMemoryInBytes"
                    memory_response = requests.get(memory_url, headers=headers, verify=False, proxies=self.PROXIES)
                    print("Device Memory Info Retrieved")
                    if memory_response.status_code == 200:
                        memory_data = memory_response.json()
                        physical_memory = memory_data.get("physicalMemoryInBytes", 0)
                        physical_memory_gb = str(int(round(physical_memory / (1024 ** 3))) if isinstance(physical_memory, (int, float)) and physical_memory != '' else '')
                        updated_values["PhysicalMemory"] = physical_memory_gb
                    else:
                        updated_values["PhysicalMemory"] = ""

                    # Retrieve user details
                    user_principal_name = device_data.get("userPrincipalName", "")
                    if user_principal_name:
                        user_url = f"https://graph.microsoft.com/v1.0/users/{user_principal_name}?$select=displayName,jobTitle,department,city,country"
                        user_response = requests.get(user_url, headers=headers, verify=False, proxies=self.PROXIES)
                        print("User Info Retrieved")
                        if user_response.status_code == 200:
                            user_data = user_response.json()
                            updated_values.update({
                                "UserDisplayName": user_data.get("displayName", ""),
                                "JobTitle": user_data.get("jobTitle", ""),
                                "Department": user_data.get("department", ""),
                                "City": user_data.get("city", ""),
                                "Country": user_data.get("country", "")
                            })

                    # Retrieve Entra device details for trustType
                    entra_device_url = f"https://graph.microsoft.com/v1.0/devices?$filter=deviceId eq '{entra_device_id}'&$select=trustType"
                    entra_response = requests.get(entra_device_url, headers=headers, verify=False, proxies=self.PROXIES)
                    print("Entra Device Info Retrieved")
                    if entra_response.status_code == 200:
                        entra_data = entra_response.json().get("value", [])
                        if entra_data:
                            updated_values["TrustType"] = entra_data[0].get("trustType", "")
                        else:
                            updated_values["TrustType"] = ""

                    # Update the database
                    self.cursor.execute('''UPDATE Devices SET 
                                            OperatingSystem = ?,
                                            OSVersion = ?,
                                            ComplianceState = ?,
                                            Manufacturer = ?,
                                            Model = ?,
                                            IntuneLastSync = ?,
                                            EntraDeviceID = ?,
                                            IntuneDeviceID = ?,
                                            UserPrincipalName = ?,
                                            MAC = ?,
                                            ReportTime = ?,
                                            TotalStorage = ?,
                                            FreeStorage = ?,
                                            PhysicalMemory = ?,
                                            UserDisplayName = ?,
                                            JobTitle = ?,
                                            Department = ?,
                                            City = ?,
                                            Country = ?,
                                            TrustType = ?
                                            WHERE DeviceName = ? AND SerialNumber = ?''',
                                        (
                                            updated_values.get("OperatingSystem", ""),
                                            updated_values.get("OSVersion", ""),
                                            updated_values.get("ComplianceState", ""),
                                            updated_values.get("Manufacturer", ""),
                                            updated_values.get("Model", ""),
                                            updated_values.get("IntuneLastSync", ""),
                                            updated_values.get("EntraDeviceID", ""),
                                            updated_values.get("IntuneDeviceID", ""),
                                            updated_values.get("UserPrincipalName", ""),
                                            updated_values.get("MAC", ""),
                                            updated_values.get("ReportTime", ""),
                                            updated_values.get("TotalStorage", ""),
                                            updated_values.get("FreeStorage", ""),
                                            updated_values.get("PhysicalMemory", ""),
                                            updated_values.get("UserDisplayName", ""),
                                            updated_values.get("JobTitle", ""),
                                            updated_values.get("Department", ""),
                                            updated_values.get("City", ""),
                                            updated_values.get("Country", ""),
                                            updated_values.get("TrustType", ""),
                                            device_name,
                                            serial_number
                                        ))
                    self.conn.commit()
                    print("Device Info Written")
                    # Update the data in the TreeView
                    for col, idx in columns_mapping.items():
                        if col in updated_values:
                            self.tree.set(selected_item, column=col, value=updated_values[col])

                    messagebox.showinfo("Success", "Device information updated successfully.")

                else:
                    messagebox.showerror("Error", f"Failed to refresh device information. Status code: {response.status_code}, Response: {response.text}")

            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while refreshing device information: {str(e)}")

    def retrieve_user_details(self, access_token, user_principal_name, device):
        # Get user details
        user_info_url = f"https://graph.microsoft.com/v1.0/users/{user_principal_name}?$select=userPrincipalName,displayName,jobTitle,department,employeeId,city,country"
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }

        try:
            # Request user details (excluding the manager field initially)
            user_info_response = requests.get(user_info_url, headers=headers, verify=False, proxies=self.PROXIES)
            if user_info_response.status_code == 200:
                user_details = user_info_response.json()

                # Add default values for missing fields
                user_details['jobTitle'] = user_details.get('jobTitle', 'N/A')
                user_details['department'] = user_details.get('department', 'N/A')
                user_details['employeeId'] = user_details.get('employeeId', 'N/A')
                user_details['city'] = user_details.get('city', 'N/A')
                user_details['country'] = user_details.get('country', 'N/A')

                # Now make a request to get the manager details using the manager endpoint
                manager_url = f"https://graph.microsoft.com/v1.0/users/{user_principal_name}/manager"
                manager_response = requests.get(manager_url, headers=headers, verify=False, proxies=self.PROXIES)
                manager_display_name = ''
                if manager_response.status_code == 200:
                    manager_info = manager_response.json()
                    manager_display_name = manager_info.get('displayName', '')

                # Add manager display name to user details
                user_details['manager_displayName'] = manager_display_name

                # Attach user details to the device
                device.update(user_details)
            else:
                print(f"Failed to retrieve user details for {user_principal_name}. Status code: {user_info_response.status_code}, Response: {user_info_response.text}")

        except Exception as e:
            print(f"Exception occurred while retrieving user details: {str(e)}")
        
    def export_to_database(self):
        # File names
        intune_file = "IntuneDeviceReport.xlsx"
        entra_file = "EntraDeviceReport.xlsx"
        user_file = "UserReport.xlsx"

        # Check if all files exist
        if not all(os.path.exists(file) for file in [intune_file, entra_file, user_file]):
            raise FileNotFoundError("One or more required files are missing.")

        # Load the reports into DataFrames
        intune_df = pd.read_excel(intune_file).fillna("")
        entra_df = pd.read_excel(entra_file).fillna("")
        user_df = pd.read_excel(user_file).fillna("")

        # Ensure proper column naming for consistency
        intune_columns = [
            "DeviceName", "UserPrincipalName", "EntraDeviceID", "IntuneDeviceID", "OperatingSystem", "OSVersion", 
            "ComplianceState", "Model", "Manufacturer", "SerialNumber", "MAC", "IntuneLastSync", "Encryption", 
            "TotalStorage", "FreeStorage", 'PhysicalMemory', "Source", "Remarks"
        ]
        entra_columns = ["DeviceName", "EntraDeviceID", "TrustType"]
        user_columns = ["UPN", "UserDisplayName", "JobTitle", "Department", "City", "Country"]

        intune_df.columns = intune_columns
        entra_df.columns = entra_columns
        user_df.columns = user_columns

        # Merge data
        merged_df = intune_df.merge(entra_df, on=["DeviceName", "EntraDeviceID"], how="left")
        merged_df = merged_df.merge(user_df, left_on="UserPrincipalName", right_on="UPN", how="left")

        # Prepare data for the Devices table
        devices_data = []
        for _, row in merged_df.iterrows():
            # Check if UserPrincipalName is empty
            if not row['UserPrincipalName']:
                user_display_name = ""
                job_title = ""
                department = ""
                city = ""
                country = ""
            else:
                user_display_name = row['UserDisplayName']
                job_title = row['JobTitle']
                department = row['Department']
                city = row['City']
                country = row['Country']
            # Convert TotalStorage, FreeStorage, and PhysicalMemory from bytes to GB
            total_storage = row['TotalStorage']
            free_storage = row['FreeStorage']

            # Convert to GB and round to the nearest integer if the value is not 'N/A'
            total_storage_gb = int(round(total_storage / (1024 ** 3))) if isinstance(total_storage, (int, float)) and total_storage != '' else ''
            free_storage_gb = int(round(free_storage / (1024 ** 3))) if isinstance(free_storage, (int, float)) and free_storage != '' else ''
            device_data = {
                'DeviceName': str(row['DeviceName']),
                'SerialNumber': str(row['SerialNumber']),
                'EntraDeviceID': str(row['EntraDeviceID']),
                'IntuneDeviceID': str(row['IntuneDeviceID']),
                'UserPrincipalName': str(row['UserPrincipalName']),
                'OperatingSystem': str(row['OperatingSystem']),
                'OSVersion': str(row['OSVersion']),
                'ComplianceState': str(row['ComplianceState']),
                'Model': str(row['Model']),
                'Manufacturer': str(row['Manufacturer']),
                'MAC': str(row['MAC']),
                'IntuneLastSync': str(row['IntuneLastSync']),
                'ReportTime': time.strftime("%Y-%m-%d %H:%M:%S"),
                'Source': str(row['Source']),
                'Encryption': str(row['Encryption']),
                'UserDisplayName': user_display_name,
                'JobTitle': job_title,
                'Department': department,
                'City': city,
                'Country': country,
                'TrustType': str(row['TrustType']),
                'TotalStorage': total_storage_gb,
                'FreeStorage': free_storage_gb,
                'PhysicalMemory': str(row['PhysicalMemory']),
                'Remarks': str(row['Remarks'])
            }

            # Skip empty records
            if not any(device_data.values()):
                continue

            devices_data.append(device_data)

        # Convert the selected device data to a DataFrame and save as an Excel file
        if devices_data:
            df = pd.DataFrame(devices_data)
            df.fillna('', inplace=True)
            df.to_excel('DeviceInventory.xlsx', index=False)

            # Write the data to the SQLite database
            with sqlite3.connect('inventory.db', timeout=30) as conn:
                cursor = conn.cursor()

                for _, row in df.iterrows():
                    # Convert ReportTime to string or None (if it's NaT)
                    report_time_str = row['ReportTime']
                    if pd.isna(report_time_str):
                        report_time_str = None
                    elif isinstance(report_time_str, pd.Timestamp):
                        report_time_str = report_time_str.strftime("%Y-%m-%d %H:%M:%S")

                    # Insert or update the record in the database
                    cursor.execute('''REPLACE INTO Devices (DeviceName, SerialNumber, UserPrincipalName, EntraDeviceID, IntuneDeviceID, OperatingSystem, 
                                        OSVersion, Source, ComplianceState, 
                                       Model, Manufacturer, MAC, IntuneLastSync, ReportTime, Encryption, UserDisplayName, JobTitle, Department, City, Country,
                                       TrustType, TotalStorage, FreeStorage, PhysicalMemory)
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                                    (row['DeviceName'], row['SerialNumber'], row['UserPrincipalName'], row['EntraDeviceID'], row['IntuneDeviceID'], 
                                    row['OperatingSystem'], row['OSVersion'], 'Cloud', row['ComplianceState'], row['Model'], 
                                    row['Manufacturer'], row['MAC'], row['IntuneLastSync'], report_time_str, row['Encryption'], row['UserDisplayName'], 
                                    row['JobTitle'], row['Department'], row['City'], row['Country'],
                                    row['TrustType'], row['TotalStorage'], row['FreeStorage'], row['PhysicalMemory']))

                conn.commit()
        else:
            print("No valid devices found to export to Excel.")

    def retrieve_bitlocker_key(self):
        selected_item = self.tree.selection()
        if selected_item:
            item_values = self.tree.item(selected_item, "values")
            # Create a mapping of column names to their index positions
            columns_mapping = {col: idx for idx, col in enumerate(self.tree["columns"])}
            device_name = str(item_values[columns_mapping["DeviceName"]])  
            serial_number = str(item_values[columns_mapping["SerialNumber"]])  
            
            # Retrieve DeviceID from the database
            self.cursor.execute("SELECT EntraDeviceID FROM Devices WHERE DeviceName = ? AND SerialNumber = ?", (device_name, serial_number))
            result = self.cursor.fetchone()
            if result:
                device_id = result[0]
            else:
                messagebox.showerror("Error", "Device ID not found for the selected device.")
                return

            # Create a new window for retrieving BitLocker Key
            key_window = tk.Toplevel(self.root)
            key_window.title("BitLocker Recovery Key")
            key_window.geometry("550x200")
            key_window.transient(self.root)
            key_window.grab_set()  # Make the main window unresponsive

            # Center the key window on the screen
            key_window.update_idletasks()
            x = (key_window.winfo_screenwidth() // 2) - (key_window.winfo_width() // 2)
            y = (key_window.winfo_screenheight() // 2) - (key_window.winfo_height() // 2)
            key_window.geometry(f"+{x}+{y}")

            def create_label_pair(parent, label_text, value_text=""):
                label_frame = tk.Frame(parent)
                label_frame.pack(anchor="w", padx=10, pady=5, fill="x")
                
                label = tk.Label(label_frame, text=label_text, font=("Arial", 10, "bold"), fg="#000000", anchor="w")
                label.pack(side="left")
                
                value_entry = tk.Entry(label_frame, font=("Arial", 10), fg="#000000", relief="flat", readonlybackground="#F0F0F0")
                value_entry.insert(0, value_text)
                value_entry.configure(state="readonly")
                value_entry.pack(side="left", fill="x", expand=True)
                
                return value_entry

            # Create labels for initial display and placeholders for values
            device_name_entry = create_label_pair(key_window, "DeviceName: ")
            device_name_entry.configure(state='normal')
            device_name_entry.delete(0, 'end')
            device_name_entry.insert(0, device_name)
            device_name_entry.configure(state='readonly')
            key_id_value_entry = create_label_pair(key_window, "Key ID: ")
            recovery_key_value_entry = create_label_pair(key_window, "Recovery Key: ")
            backup_time_value_entry = create_label_pair(key_window, "Backup Time: ")

            def retrieve_key():
                try:
                    # Retrieve all BitLocker recovery keys from Microsoft Graph API
                    access_token = self.get_access_token()
                    headers = {
                        'Authorization': f'Bearer {access_token}',
                        'Content-Type': 'application/json'
                    }
                    url = 'https://graph.microsoft.com/v1.0/informationProtection/bitlocker/recoveryKeys'
                    response = requests.get(url, headers=headers, verify=False, proxies=self.PROXIES)

                    if response.status_code == 200:
                        recovery_keys = response.json().get('value', [])
                        bitlocker_key_id = None
                        recovery_key = None
                        backup_time = None

                        # Find the corresponding BitLocker Key ID using the Device ID
                        for key in recovery_keys:
                            if key.get('deviceId', '').lower() == device_id.lower():
                                bitlocker_key_id = key.get('id')
                                break

                        if bitlocker_key_id:
                            # Retrieve the actual recovery key using the BitLocker Key ID
                            key_url = f'https://graph.microsoft.com/v1.0/informationProtection/bitlocker/recoveryKeys/{bitlocker_key_id}?$select=key'
                            key_response = requests.get(key_url, headers=headers, verify=False, proxies=self.PROXIES)

                            if key_response.status_code == 200:
                                key_data = key_response.json()
                                recovery_key = key_data.get('key')
                                backup_time = key_data.get('createdDateTime', 'N/A')

                                with sqlite3.connect('inventory.db', timeout=30) as conn:
                                    cursor = conn.cursor()
                                    cursor.execute('''REPLACE INTO BitLockerKeys (DeviceName, SerialNumber, KeyID, RecoveryKey, BackupTime) 
                                                    VALUES (?, ?, ?, ?, ?)''', 
                                                (device_name, serial_number, bitlocker_key_id, recovery_key, backup_time))
                                    conn.commit()

                                messagebox.showinfo("Retrieved Key", "Recovery key retrieved successfully.")
                                # Update the labels with the retrieved values
                                key_id_value_entry.configure(state='normal')
                                key_id_value_entry.delete(0, 'end')
                                key_id_value_entry.insert(0, bitlocker_key_id)
                                key_id_value_entry.configure(state='readonly')

                                recovery_key_value_entry.configure(state='normal')
                                recovery_key_value_entry.delete(0, 'end')
                                recovery_key_value_entry.insert(0, recovery_key)
                                recovery_key_value_entry.configure(state='readonly')

                                backup_time_value_entry.configure(state='normal')
                                backup_time_value_entry.delete(0, 'end')
                                backup_time_value_entry.insert(0, backup_time)
                                backup_time_value_entry.configure(state='readonly')
                            else:
                                self.retrieve_local_key(device_name, serial_number, key_id_value_entry, recovery_key_value_entry, backup_time_value_entry)
                        else:
                            self.retrieve_local_key(device_name, serial_number, key_id_value_entry, recovery_key_value_entry, backup_time_value_entry)
                    else:
                        self.retrieve_local_key(device_name, serial_number, key_id_value_entry, recovery_key_value_entry, backup_time_value_entry)
                except Exception as e:
                    messagebox.showerror("Error", "No BitLocker Key ID found for this device.")        

            retrieve_button = tk.Button(key_window, text="Retrieve Key", command=retrieve_key, font=("Arial", 10), bg="#4A5568", fg="#E4E7EB", activebackground="#6B7280")
            retrieve_button.pack(pady=10)

    def retrieve_local_key(self, device_name, serial_number, key_id_value_entry, recovery_key_value_entry, backup_time_value_entry):
        with sqlite3.connect('inventory.db', timeout=30) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT KeyID, RecoveryKey, BackupTime FROM BitLockerKeys WHERE DeviceName = ? AND SerialNumber = ?",
                        (device_name, serial_number))
            existing_key = cursor.fetchone()
            if existing_key:
                key_id, recovery_key, backup_time = existing_key
                messagebox.showinfo("Retrieved Key", "Recovery key retrieved from local database.")

                # Update the labels with the existing values
                key_id_value_entry.configure(state='normal')
                key_id_value_entry.delete(0, 'end')
                key_id_value_entry.insert(0, key_id)
                key_id_value_entry.configure(state='readonly')

                recovery_key_value_entry.configure(state='normal')
                recovery_key_value_entry.delete(0, 'end')
                recovery_key_value_entry.insert(0, recovery_key)
                recovery_key_value_entry.configure(state='readonly')

                backup_time_value_entry.configure(state='normal')
                backup_time_value_entry.delete(0, 'end')
                backup_time_value_entry.insert(0, backup_time)
                backup_time_value_entry.configure(state='readonly')
            else:
                messagebox.showerror("Error", "No BitLocker Key ID found for this device.")

    def download_template(self):
        # Prompt user to choose save location for the template
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Device inventory template", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                # Retrieve the headers from the Devices table
                columns = [
                    'DeviceName', 'SerialNumber', 'OperatingSystem', 'OSVersion', 'Manufacturer', 'Model', 'TotalStorage', 
                    'PhysicalMemory', 'UserPrincipalName', 'UserDisplayName', 'Department', 'Country', 'City'
                ]
                # Create an empty DataFrame with the column headers
                df = pd.DataFrame(columns=columns)
                # Write the DataFrame to the Excel file
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Success", "Template downloaded successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to download template: {e}")

    def upload_data(self):
        # Prompt user to choose an Excel file to upload
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                # Load data from the selected Excel file
                new_data = pd.read_excel(file_path).fillna("")

                # Check for illegal columns
                expected_columns = [
                    'DeviceName', 'SerialNumber', 'OperatingSystem', 'OSVersion', 'Manufacturer', 'Model', 'TotalStorage', 
                    'PhysicalMemory', 'UserPrincipalName', 'UserDisplayName', 'Department', 'Country', 'City'
                ]
                illegal_columns = [col for col in new_data.columns if col not in expected_columns]
                if illegal_columns:
                    messagebox.showerror("Error", f"Illegal column(s) detected. Please remove them before uploading the file!")
                    return

                # Check for missing DeviceName or SerialNumber
                for index, row in new_data.iterrows():
                    if pd.isna(row['DeviceName']) or pd.isna(row['SerialNumber']) or row['DeviceName'] == "" or row['SerialNumber'] == "":
                        messagebox.showerror("Error", f"Row {index + 2} is missing DeviceName and/or SerialNumber.")
                        return

                new_data.dropna(subset=['DeviceName', 'SerialNumber'], how='all', inplace=True)

                # Automatically add the current timestamp for ReportTime
                current_time = time.strftime("%Y-%m-%d %H:%M:%S")
                new_data["ReportTime"] = current_time

                # Insert data into Devices table
                with sqlite3.connect('inventory.db', timeout=30) as conn:
                    cursor = conn.cursor()

                    for _, row in new_data.iterrows():                                               
                        # Insert or replace the record into Devices table
                        cursor.execute('''REPLACE INTO Devices (DeviceName, SerialNumber, ReportTime, 'OperatingSystem', 'OSVersion', 
                                       'Manufacturer', 'Model', 'TotalStorage', 
                                        'PhysicalMemory', 'UserPrincipalName', 'UserDisplayName', 'Department', 'Country', 'City')
                                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                                        (row['DeviceName'], row['SerialNumber'], row['ReportTime'], row['OperatingSystem'], row['OSVersion'],
                                         row['Manufacturer'], row['Model'], row['TotalStorage'], row['PhysicalMemory'], row['UserPrincipalName'],
                                         row['UserDisplayName'], row['Department'], row['Country'], row['City']))
                        
                        if not pd.isna(row['PhysicalMemory']) and not row['PhysicalMemory'] == "":
                            cursor.execute('''REPLACE INTO Memory (DeviceName, SerialNumber, 'PhysicalMemory')
                                            VALUES (?, ?, ?)''',
                                            (row['DeviceName'], row['SerialNumber'], row['PhysicalMemory']))

                    conn.commit()
                messagebox.showinfo("Success", "Data uploaded successfully.") 
                self.load_data()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to upload data: {e}")

    def main_app(self, inventory_type):
        if inventory_type == 'software':
            # Create a menu bar
            self.menubar = tk.Menu(self.root, bg="#1F2933", fg="#E4E7EB")
            self.root.config(menu=self.menubar)

            # Create File menu
            self.file_menu = tk.Menu(self.menubar, tearoff=0)
            self.menubar.add_cascade(label="File", menu=self.file_menu)
                    
            self.file_menu.add_command(label="Refresh Data", command=self.refresh_software_data)
            self.file_menu.add_command(label="Exit", command=self.root.quit)

            # Create Data menu
            self.data_menu = tk.Menu(self.menubar, tearoff=0)
            self.menubar.add_cascade(label="Data", menu=self.data_menu)
            
            # Create Export sub-menu
            self.export_menu = tk.Menu(self.data_menu, tearoff=0)
            self.data_menu.add_cascade(label="Export", menu=self.export_menu)
            self.export_menu.add_command(label="Export View", command=self.export_software_view)
            self.export_menu.add_command(label="Export Database", command=self.export_software_database)

            # Add Return option at the end of the menu
            self.menubar.add_command(label="Return", command=self.confirm_return)

            # Set default padding and window resizing behavior
            self.root.configure(padx=10, pady=10, bg="#323F4B")
            self.root.columnconfigure(0, weight=1)
            self.root.rowconfigure(1, weight=1)

            self.search_filter_frame = tk.Frame(self.root, padx=10, pady=5, bg="#3E4C59")
            self.search_filter_frame.pack(fill='x', pady=5)
            
            # Split the filter panel into two halves
            self.left_panel = tk.Frame(self.search_filter_frame, bg="#3E4C59")
            self.left_panel.grid(row=0, column=0, sticky="nsew")
            
            self.right_panel = tk.Frame(self.search_filter_frame, bg="#3E4C59")
            self.right_panel.grid(row=0, column=1, sticky="nsew")
            
            self.search_filter_frame.columnconfigure(0, weight=3)
            self.search_filter_frame.columnconfigure(1, weight=1)

            # Updated the total records label to ensure the position remains fixed
            self.total_records_panel = tk.Frame(self.left_panel, bg="#3E4C59")
            self.total_records_panel.grid(row=0, column=4, columnspan=2, sticky="nsew")
            self.total_records_label = tk.Label(self.total_records_panel, text="Records Displayed:", font=("Arial", 9, "bold"), fg="#E4E7EB", bg="#3E4C59", anchor="c")
            self.total_records_displayed_label = tk.Label(self.total_records_panel, text="000000", font=("Arial", 9, "bold"), fg="#66CDAA", bg="#3E4C59", anchor="c")
            self.total_records_total_label = tk.Label(self.total_records_panel, text="/000000", font=("Arial", 9, "bold"), fg="#66CDAA", bg="#3E4C59", anchor="c")

            self.update_total_records_label(0, 0)
            self.total_records_label.grid(row=0, column=0, padx=0, pady=5, sticky="ew")
            self.total_records_displayed_label.grid(row=0, column=1, padx=5, pady=5, sticky="e")
            self.total_records_total_label.grid(row=0, column=2, padx=5, pady=5, sticky="e")

            self.software_data = pd.DataFrame()  # Initialize data attribute
            self.software_sort_orders = {}  # Track sort order for each column
            self.min_window_width = 1200  # Default minimum window width
            self.software_filtered_data = pd.DataFrame()  # Initialize filtered_data attribute

            # Search UI
            self.search_label = tk.Label(self.left_panel, text="Search:", font=("Arial", 10, "bold"), fg="#E4E7EB", bg="#3E4C59")
            self.search_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
            self.software_search_entry = tk.Entry(self.left_panel, font=("Arial", 10), bg="#F0F0F0", fg="#000000")
            self.software_search_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")            

            self.search_filter_frame.columnconfigure(1, weight=1)

            # Apply and Clear buttons
            self.software_apply_button = tk.Button(self.left_panel, text="Search", command=self.apply_software_filter, font=("Arial", 9), bg="#4A5568", fg="#E4E7EB", activebackground="#6B7280")
            self.software_apply_button.grid(row=0, column=2, pady=5, padx=5, sticky="ew")
            self.software_clear_search_button = tk.Button(self.left_panel, text="Clear", command=self.clear_software_filter, font=("Arial", 9), bg="#4A5568", fg="#E4E7EB", activebackground="#6B7280")
            self.software_clear_search_button.grid(row=0, column=3, padx=5, pady=5, sticky="w")

            # Treeview to display data
            self.softare_tree_frame = tk.Frame(self.root, bg="#3E4C59")
            self.softare_tree_frame.pack(expand=True, fill='both', padx=10, pady=10)

            self.software_tree_scroll = ttk.Scrollbar(self.softare_tree_frame, orient="vertical")
            self.software_tree_scroll.pack(side="right", fill="y")

            self.software_tree_scroll_x = ttk.Scrollbar(self.softare_tree_frame, orient="horizontal")
            self.software_tree_scroll_x.pack(side="bottom", fill="x")

            self.software_tree = ttk.Treeview(self.softare_tree_frame, yscrollcommand=self.software_tree_scroll.set, xscrollcommand=self.software_tree_scroll_x.set, style="Custom.Treeview")
            self.software_tree.pack(side='left', expand=True, fill='both')

            self.software_tree_scroll.config(command=self.software_tree.yview)
            self.software_tree_scroll_x.config(command=self.software_tree.xview)

            # Bind right-click menu for copying cell value and adding remark
            self.software_tree.bind("<Button-3>", self.show_software_context_menu)

            self.software_tree["columns"] = ["SoftwareName", "Version", "InstalledDevices"]
            self.software_tree["show"] = "headings"

            self.software_tree.heading("SoftwareName", text="Software Name")
            self.software_tree.heading("Version", text="Version")
            self.software_tree.heading("InstalledDevices", text="Installed Devices")

            self.software_tree.column("SoftwareName", width=300, anchor="w")
            self.software_tree.column("Version", width=150, anchor="center")
            self.software_tree.column("InstalledDevices", width=150, anchor="center")

            self.software_tree.pack(expand=True, fill="both")

            # Context menu for right-click
            self.software_context_menu = tk.Menu(self.software_tree, tearoff=0)
            self.software_context_menu.add_command(label="Copy", command=self.copy_software_cell_value)

            # Style for Treeview with borders
            style = ttk.Style()
            style.configure("Custom.Treeview", rowheight=25, borderwidth=1, relief="solid", font=("Arial", 10), foreground="#E4E7EB", background="#323F4B", fieldbackground="#323F4B")
            style.configure("Treeview.Heading", font=("Arial", 10, "bold"), foreground="#3E4C59", background="#1F2933")
            style.layout("Custom.Treeview", [
                ("Custom.Treeview.treearea", {'sticky': 'nswe'})
            ])

            # Pagination Frame for the buttons
            self.pagination_frame = tk.Frame(self.root, bg="#3E4C59", pady=10)
            self.pagination_frame.pack(fill='x', pady=5)

            # Pagination Buttons
            self.software_previous_button = tk.Button(self.pagination_frame, text="Previous Page", command=self.previous_software_page, font=("Arial", 10), bg="#4A5568", fg="#E4E7EB", activebackground="#6B7280")
            self.software_previous_button.pack(side='left', padx=10)

            # Page Number Label (moved to the center)
            self.softwarepage_number_label = tk.Label(self.pagination_frame, text="Page 1 / 1", font=("Arial", 10), bg="#3E4C59", fg="#E4E7EB")
            self.softwarepage_number_label.pack(side='left', expand=True)

            self.software_next_button = tk.Button(self.pagination_frame, text="Next Page", command=self.next_software_page, font=("Arial", 10), bg="#4A5568", fg="#E4E7EB", activebackground="#6B7280")
            self.software_next_button.pack(side='right', padx=10)
            
            # Set minimum window width based on total column width
            total_width = max(self.min_window_width, sum([max(100, len(col) * 10) for col in self.software_tree['columns']]))
            self.root.minsize(width=total_width, height=600)
            self.root.geometry(f"{total_width}x600") 

            # Load Excel data automatically
            self.load_software_data()

            # Update the display
            self.root.update()

        else:
            # Create a menu bar
            self.menubar = tk.Menu(self.root, bg="#1F2933", fg="#E4E7EB")
            self.root.config(menu=self.menubar)

            # Create File menu
            self.file_menu = tk.Menu(self.menubar, tearoff=0)
            self.menubar.add_cascade(label="File", menu=self.file_menu)
                    
            self.file_menu.add_command(label="Refresh Data", command=self.refresh_data)
            self.file_menu.add_command(label="Exit", command=self.root.quit)

            # Create View menu
            self.view_menu = tk.Menu(self.menubar, tearoff=0)
            self.menubar.add_cascade(label="View", menu=self.view_menu)

            # Create Columns sub-menu under View menu
            self.columns_menu = tk.Menu(self.view_menu, tearoff=0)
            self.view_menu.add_cascade(label="Columns", menu=self.columns_menu)

            # Get columns from Devices table and add as sub-options under Columns (excluding certain columns)
            self.columns_to_exclude = ["DeviceName", "SerialNumber", "OperatingSystem", "IntuneLastSync", "Source", "Remarks"]
            with sqlite3.connect('inventory.db', timeout=30) as conn:
                cursor = conn.cursor()
                cursor.execute("PRAGMA table_info(Devices)")
                columns = [info[1] for info in cursor.fetchall() if info[1] not in self.columns_to_exclude]

            self.column_vars = {}
            for column in columns:
                self.column_vars[column] = tk.BooleanVar(value=False)
                self.columns_menu.add_checkbutton(
                    label=column,
                    onvalue=True,
                    offvalue=False,
                    variable=self.column_vars[column],
                    command=lambda col=column: self.toggle_column(col)
                )

            # Initialize selected columns
            self.selected_columns = self.columns_to_exclude.copy()

            # Create Data menu
            self.data_menu = tk.Menu(self.menubar, tearoff=0)
            self.menubar.add_cascade(label="Data", menu=self.data_menu)
            
            # Create Import sub-menu
            self.import_menu = tk.Menu(self.data_menu, tearoff=0)
            self.data_menu.add_cascade(label="Import", menu=self.import_menu)
            self.import_menu.add_command(label="Download Template", command=self.download_template)
            self.import_menu.add_command(label="Import Data", command=self.upload_data)

            # Create Export sub-menu
            self.export_menu = tk.Menu(self.data_menu, tearoff=0)
            self.data_menu.add_cascade(label="Export", menu=self.export_menu)
            self.export_menu.add_command(label="Export View", command=self.export_view)
            self.export_menu.add_command(label="Export Database", command=self.export_database)

            # Add Return option at the end of the menu
            self.menubar.add_command(label="Return", command=self.confirm_return)

            # Set default padding and window resizing behavior
            self.root.configure(padx=10, pady=10, bg="#323F4B")
            self.root.columnconfigure(0, weight=1)
            self.root.rowconfigure(1, weight=1)

            self.search_filter_frame = tk.Frame(self.root, padx=10, pady=5, bg="#3E4C59")
            self.search_filter_frame.pack(fill='x', pady=5)
            
            # Split the filter panel into two halves
            self.left_panel = tk.Frame(self.search_filter_frame, bg="#3E4C59")
            self.left_panel.grid(row=0, column=0, sticky="nsew")
            
            self.right_panel = tk.Frame(self.search_filter_frame, bg="#3E4C59")
            self.right_panel.grid(row=0, column=1, sticky="nsew")
            
            self.search_filter_frame.columnconfigure(0, weight=3)
            self.search_filter_frame.columnconfigure(1, weight=1)

            # Updated the total records label to ensure the position remains fixed
            self.total_records_panel = tk.Frame(self.left_panel, bg="#3E4C59")
            self.total_records_panel.grid(row=0, column=3, columnspan=2, sticky="nsew")
            self.total_records_label = tk.Label(self.total_records_panel, text="Records Displayed:", font=("Arial", 9, "bold"), fg="#E4E7EB", bg="#3E4C59", anchor="c")
            self.total_records_displayed_label = tk.Label(self.total_records_panel, text="000000", font=("Arial", 9, "bold"), fg="#66CDAA", bg="#3E4C59", anchor="c")
            self.total_records_total_label = tk.Label(self.total_records_panel, text="/000000", font=("Arial", 9, "bold"), fg="#66CDAA", bg="#3E4C59", anchor="c")

            self.update_total_records_label(0, 0)
            self.total_records_label.grid(row=0, column=0, padx=0, pady=5, sticky="ew")
            self.total_records_displayed_label.grid(row=0, column=1, padx=5, pady=5, sticky="e")
            self.total_records_total_label.grid(row=0, column=2, padx=5, pady=5, sticky="e")

            self.data = pd.DataFrame()  # Initialize data attribute
            self.sort_orders = {}  # Track sort order for each column
            self.min_window_width = 1200  # Default minimum window width
            self.filters = {}  # Track active filters
            self.filtered_data = pd.DataFrame()  # Initialize filtered_data attribute

            # Search UI
            self.search_label = tk.Label(self.left_panel, text="Search:", font=("Arial", 10, "bold"), fg="#E4E7EB", bg="#3E4C59")
            self.search_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
            self.search_entry = tk.Entry(self.left_panel, font=("Arial", 10), bg="#F0F0F0", fg="#000000")
            self.search_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
            self.clear_search_button = tk.Button(self.left_panel, text="Clear Search", command=self.clear_search_entry, font=("Arial", 9), bg="#4A5568", fg="#E4E7EB", activebackground="#6B7280")
            self.clear_search_button.grid(row=0, column=2, padx=5, pady=5, sticky="w")

            self.search_filter_frame.columnconfigure(1, weight=1)

            # Filter UI for two filters with operator
            self.filter_labels = []
            self.filter_column_comboboxes = []
            self.filter_operator_comboboxes = []
            self.filter_value_entries = []

            for i in range(3):  # Creating 3 filters
                row = i + 1
                # Filter Label
                filter_label = tk.Label(self.left_panel, text=f"Filter {i + 1}:", font=("Arial", 10, "bold"), fg="#E4E7EB", bg="#3E4C59")
                filter_label.grid(row=row, column=0, padx=5, pady=5, sticky="w")  # Consistent padding
                self.filter_labels.append(filter_label)

                # Filter Column Combobox
                filter_column_combobox = ttk.Combobox(self.left_panel, values=["", "DeviceName", "SerialNumber", "OperatingSystem", "Source", "Remarks"], state="readonly")
                filter_column_combobox.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
                self.filter_column_comboboxes.append(filter_column_combobox)

                # Filter Operator Combobox
                filter_operator_combobox = ttk.Combobox(self.left_panel, values=["equals", "does not equal", "contains", "does not contain", "begins with", "does not begin with", "ends with", "does not end with"], state="readonly")
                filter_operator_combobox.set("equals")
                filter_operator_combobox.grid(row=row, column=2, padx=5, pady=5, sticky="ew")
                self.filter_operator_comboboxes.append(filter_operator_combobox)

                # Filter Value Entry
                filter_value_entry = tk.Entry(self.left_panel, font=("Arial", 10), bg="#F0F0F0", fg="#000000")
                filter_value_entry.grid(row=row, column=3, padx=5, pady=5, sticky="ew")
                self.filter_value_entries.append(filter_value_entry)

            # Global Operator Combobox
            self.operator_label = tk.Label(self.left_panel, text="Operator:", font=("Arial", 10, "bold"), fg="#E4E7EB", bg="#3E4C59")
            self.operator_label.grid(row=1, column=4, padx=5, pady=5, sticky="w")
            self.operator_combobox = ttk.Combobox(self.left_panel, values=["AND", "OR"], state="readonly")
            self.operator_combobox.set("AND")
            self.operator_combobox.grid(row=1, column=5, padx=5, pady=5, sticky="ew")

            # Apply and Clear buttons
            self.apply_button = tk.Button(self.left_panel, text="Apply", command=self.apply_filters, font=("Arial", 9), bg="#4A5568", fg="#E4E7EB", activebackground="#6B7280")
            self.apply_button.grid(row=2, column=4, columnspan=2, pady=5, padx=5, sticky="ew")

            self.clear_button = tk.Button(self.left_panel, text="Clear", command=self.clear_filters, font=("Arial", 9), bg="#4A5568", fg="#E4E7EB", activebackground="#6B7280")
            self.clear_button.grid(row=3, column=4, columnspan=2, pady=5, padx=5, sticky="ew")

            # Treeview to display data
            self.tree_frame = tk.Frame(self.root, bg="#3E4C59")
            self.tree_frame.pack(expand=True, fill='both', padx=10, pady=10)

            self.tree_scroll = ttk.Scrollbar(self.tree_frame, orient="vertical")
            self.tree_scroll.pack(side="right", fill="y")

            self.tree_scroll_x = ttk.Scrollbar(self.tree_frame, orient="horizontal")
            self.tree_scroll_x.pack(side="bottom", fill="x")

            self.tree = ttk.Treeview(self.tree_frame, yscrollcommand=self.tree_scroll.set, xscrollcommand=self.tree_scroll_x.set, style="Custom.Treeview")
            self.tree.pack(side='left', expand=True, fill='both')

            self.tree_scroll.config(command=self.tree.yview)
            self.tree_scroll_x.config(command=self.tree.xview)

            # Bind right-click menu for copying cell value and adding remark
            self.tree.bind("<Button-3>", self.show_context_menu)

            # Context menu for right-click
            self.context_menu = tk.Menu(self.tree, tearoff=0)
            self.context_menu.add_command(label="Copy", command=self.copy_cell_value)
            self.context_menu.add_command(label="Add/Edit Remark", command=self.add_edit_remark)
            self.context_menu.add_command(label="BitLocker Keys", command=self.retrieve_bitlocker_key)
            self.context_menu.add_command(label="Refresh", command=self.refresh_device_info)

            # Style for Treeview with borders
            style = ttk.Style()
            style.configure("Custom.Treeview", rowheight=25, borderwidth=1, relief="solid", font=("Arial", 10), foreground="#E4E7EB", background="#323F4B", fieldbackground="#323F4B")
            style.configure("Treeview.Heading", font=("Arial", 10, "bold"), foreground="#3E4C59", background="#1F2933")
            style.layout("Custom.Treeview", [
                ("Custom.Treeview.treearea", {'sticky': 'nswe'})
            ])

            # Pagination Frame for the buttons
            self.pagination_frame = tk.Frame(self.root, bg="#3E4C59", pady=10)
            self.pagination_frame.pack(fill='x', pady=5)

            # Pagination Buttons
            self.previous_button = tk.Button(self.pagination_frame, text="Previous Page", command=self.previous_page, font=("Arial", 10), bg="#4A5568", fg="#E4E7EB", activebackground="#6B7280")
            self.previous_button.pack(side='left', padx=10)

            # Page Number Label (moved to the center)
            self.page_number_label = tk.Label(self.pagination_frame, text="Page 1 / 1", font=("Arial", 10), bg="#3E4C59", fg="#E4E7EB")
            self.page_number_label.pack(side='left', expand=True)

            self.next_button = tk.Button(self.pagination_frame, text="Next Page", command=self.next_page, font=("Arial", 10), bg="#4A5568", fg="#E4E7EB", activebackground="#6B7280")
            self.next_button.pack(side='right', padx=10)
            
            # Set minimum window width based on total column width
            total_width = max(self.min_window_width, sum([max(100, len(col) * 10) for col in self.tree['columns']]))
            self.root.minsize(width=total_width, height=600)
            self.root.geometry(f"{total_width}x600") 

            # Load Excel data automatically
            self.load_data()

            # Update the display
            self.root.update()

    def export_view(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Export View"
        )
        if file_path:
            try:
                # Export the filtered data to an Excel file
                self.filtered_data.to_excel(file_path, index=False)
                messagebox.showinfo("Success", "Filtered view exported successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export view: {e}")
    
    def export_database(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Export Database"
        )
        if file_path:
            try:
                # Export the entire database to an Excel file
                with sqlite3.connect('inventory.db', timeout=30) as conn:
                    data = pd.read_sql_query("SELECT * FROM Devices", conn)
                    data.to_excel(file_path, index=False)
                messagebox.showinfo("Success", "Database exported successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export database: {e}")

    def toggle_column(self, column):
        if column in self.selected_columns:
            # If the column is already selected, remove it
            self.selected_columns.remove(column)
        else:
            # If the column is not selected, add it before the 'Source' column
            if "Source" in self.selected_columns:
                source_index = self.selected_columns.index("Source")
                self.selected_columns.insert(source_index, column)
            else:
                # If 'Source' is not in the selected columns, just add it to the end
                self.selected_columns.append(column)

        # Update the Treeview with selected columns
        self.update_treeview_columns()

    def adjust_window_width(self):
        # Calculate the total width required based on the current columns
        total_width = sum(max(100, len(col) * 10) for col in self.selected_columns)  # Default to 100 width or header length * 10
        total_width = max(self.min_window_width, total_width)  # Ensure the width is at least the minimum window width

        # Set the new window width while maintaining height
        current_height = self.root.winfo_height()
        self.root.geometry(f"{total_width}x{current_height}")
        self.root.minsize(width=total_width, height=current_height)

        # Update the display to reflect changes
        self.root.update_idletasks()

    def on_window_configure(self, event):
        # Allow resizing to avoid locking the window after maximizing/unmaximizing
        if self.root.state() == "normal" and event.widget == self.root:
            current_width = self.root.winfo_width()
            current_height = self.root.winfo_height()

            # Reset height to a reasonable value if it remains too large after unmaximizing
            if current_height > 800:  # Arbitrary value to detect maximized height
                self.root.geometry(f"{current_width}x600")  # Restore to a reasonable height

        # Ensure the minimum height remains at 600
        self.root.minsize(width=400, height=600)  # Set a minimum width and height

    def update_treeview_columns(self, is_new_column=False):
        # Define the new column width equivalent to 10 characters
        default_char_width = 10  # Estimated width of a single character in pixels
        new_column_width = default_char_width * 10

        # If columns are being added, update the UI width
        if is_new_column:
            if hasattr(self, "initial_total_width"):
                self.initial_total_width += new_column_width
            else:
                self.initial_total_width = max(self.root.winfo_width(), self.min_window_width)

            # Ensure the total width does not exceed the maximum allowed width (1310 pixels)
            self.total_width = min(self.initial_total_width, 1310)
            self.root.minsize(width=self.total_width, height=600)

        current_columns = list(self.tree["columns"])
        current_column_widths = {col: self.tree.column(col, option="width") for col in current_columns}

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.tree["column"] = self.selected_columns
        self.tree["show"] = "headings"

        # Synchronize the menu checkboxes with the selected columns
        for col, var in self.column_vars.items():
            var.set(col in self.selected_columns)

        # Configure TreeView headings and column properties
        for col in self.selected_columns:
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_data(_col))
            header_width = max(100, len(col) * 10)  # Ensure a default minimum width for headers
            self.tree.column(col, width=header_width, minwidth=header_width, stretch=True, anchor="center")

        # Insert the data into the TreeView
        for _, row in self.filtered_data.iterrows():
            row_values = [row[col] for col in self.selected_columns]
            self.tree.insert("", "end", values=row_values)

        # Adjust window width to match the selected columns
        self.adjust_window_width()

        # Update filter comboboxes with the new list of columns
        self.update_filter_columns()

        # Refresh the display
        self.root.update_idletasks()

    def update_filter_columns(self):
        """
        Update the filter comboboxes with the current list of columns in the TreeView.
        """
        # Define columns to exclude from filtering
        columns_to_exclude = ['SerialNumber', 'IntuneDeviceID', 'EntraDeviceID', 'MAC', 'ReportTime', 'IntuneLastSync']

        # Create the list of available columns for filtering
        available_columns = [col for col in self.selected_columns if col not in columns_to_exclude]

        # Update each filter column combobox with the available columns
        for combobox in self.filter_column_comboboxes:
            combobox["values"] = [""] + available_columns
            if combobox.get() not in available_columns:
                combobox.set("")  # Reset combobox if the current value is not valid

    def clear_search_and_filters_inputs(self):
        # Clear the search entry field
        self.search_entry.delete(0, tk.END)

        # Reset filter comboboxes
        for i, combobox in enumerate(self.filter_column_comboboxes):
            if isinstance(combobox, ttk.Combobox):  # Ensure it's a Combobox
                combobox.set("")  # Clear Comboboxes (filter column selectors)
            else:
                print(f"Unexpected widget in filter_column_comboboxes at index {i}: {type(combobox)}")

        # Reset value entries (handles both Entry and Combobox types)
        for i, entry in enumerate(self.filter_value_entries):
            if isinstance(entry, ttk.Combobox):  # If it's a Combobox, reset it
                entry.set("")
            elif isinstance(entry, tk.Entry):  # If it's an Entry widget, clear it
                entry.delete(0, tk.END)
                entry.configure(state="normal")  # Enable Entry widget in case it's disabled
            else:
                print(f"Unexpected widget in filter_value_entries at index {i}: {type(entry)}")

        # Reset the logical operator combobox
        self.operator_combobox.set("AND")

    def force_redraw(self):
        # Trigger a sort operation on the first column to force Treeview to adjust
        if self.selected_columns:
            first_column = self.selected_columns[0]
            self.sort_data(first_column)

    def confirm_return(self):
        # Prompt user for confirmation before returning to the main page
        confirm = messagebox.askyesno("Confirm Return", "Are you sure you want to return to the main page?")
        if confirm:
            # Reset the selected columns to default
            self.selected_columns = self.columns_to_exclude.copy()

            # Uncheck all menu options for non-default columns
            for col, var in self.column_vars.items():  # Assuming column_vars is a dict of {column_name: BooleanVar}
                if col not in self.columns_to_exclude:
                    var.set(False)

            # Update TreeView to reflect only default columns
            self.update_treeview_columns()

            # Return to the main screen
            self.return_to_main_screen()

    def refresh_data(self):
        # Clear all search, filter, and operator inputs before refreshing
        self.clear_search_and_filters_inputs()

        # Create a new window for the progress bar
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Refreshing Data")
        progress_window.geometry("300x100")
        progress_window.transient(self.root)
        progress_window.grab_set()  # Make the main window unresponsive

        # Center the progress window on the screen
        progress_window.update_idletasks()
        x = (progress_window.winfo_screenwidth() // 2) - (progress_window.winfo_width() // 2)
        y = (progress_window.winfo_screenheight() // 2) - (progress_window.winfo_height() // 2)
        progress_window.geometry(f"+{x}+{y}")

        # Add a progress bar widget
        progress = ttk.Progressbar(progress_window, orient="horizontal", length=250, mode="indeterminate")
        progress.pack(pady=20)
        progress.start()

        # Run the refresh in a separate thread to keep the UI responsive
        threading.Thread(target=self.run_refresh_data, args=(progress_window, progress, None)).start()

    def clear_search_and_filters_inputs(self):

        # Clear the search entry field
        self.search_entry.delete(0, tk.END)

        # Reset filter column comboboxes
        for i, combobox in enumerate(self.filter_column_comboboxes):
            if isinstance(combobox, ttk.Combobox):  # Ensure it's a Combobox
                combobox.set("")  # Clear Combobox values
            else:
                print(f"Unexpected widget in filter_column_comboboxes at index {i}: {type(combobox)}")

        # Clear filter value entries (tk.Entry widgets)
        for i, entry in enumerate(self.filter_value_entries):
            if isinstance(entry, tk.Entry):  # Ensure it's an Entry widget
                entry.delete(0, tk.END)  # Clear the text
                entry.configure(state="normal")  # Ensure Entry is enabled
            else:
                print(f"Unexpected widget in filter_value_entries at index {i}: {type(entry)}")

        # Reset the logical operator combobox
        self.operator_combobox.set("AND")

    def run_refresh_data(self, progress_window, progress, *args):
        file_path = r"DeviceInventory.xlsx"
        try:
            # Export the retrieved device information to the database
            self.export_to_database()                
            new_data = pd.read_excel(file_path).fillna("")  # Ensure no NaN values
            #new_data.dropna(subset=['DeviceName', 'SerialNumber'], how='all', inplace=True)

            # Create a new SQLite connection for this thread
            with sqlite3.connect('inventory.db', timeout=10) as thread_conn:
                thread_cursor = thread_conn.cursor()

                # Remove exact duplicate rows from new data to avoid redundant checks
                new_data = new_data.drop_duplicates()

                # Iterate through new data to check for duplicates or updates
                latest_records = set((str(row['SerialNumber']).strip().lower(), str(row['DeviceName']).strip().lower()) for _, row in new_data.iterrows())

                for _, new_row in new_data.iterrows():
                    serial_number = new_row['SerialNumber']
                    device_name = new_row['DeviceName']
                    thread_cursor.execute('SELECT * FROM Devices WHERE DeviceName = ? AND SerialNumber = ?', (device_name, serial_number,))
                    existing_record = thread_cursor.fetchone()

                    # Convert ReportTime to string without microseconds if it is a pandas.Timestamp
                    if pd.isnull(new_row['ReportTime']):
                        report_time_str = ""
                    elif isinstance(new_row['ReportTime'], pd.Timestamp):
                        report_time_str = new_row['ReportTime'].strftime("%Y-%m-%d %H:%M:%S")
                    else:
                        report_time_str = str(new_row['ReportTime'])

                    if existing_record:
                        existing_date = pd.to_datetime(existing_record[-1], errors='coerce')
                        new_date = pd.to_datetime(new_row['ReportTime'], errors='coerce')

                        # Update the existing record and set Source to "Cloud" if the record is found in the latest data
                        source_value = 'Cloud' if (str(new_row['SerialNumber']).strip().lower(), str(new_row['DeviceName']).strip().lower()) in latest_records else 'Local'

                        if new_date > existing_date:
                            thread_cursor.execute('''UPDATE Devices SET 
                                                    UserPrincipalName = ?, OperatingSystem = ?, EntraDeviceID = ?, IntuneDeviceID = ?, OSVersion = ?, 
                                                    ComplianceState = ?, Source = ?,
                                                    Model = ?, Manufacturer = ?, MAC = ?, IntuneLastSync = ?, ReportTime = ?, Encryption = ?,
                                                    UserDisplayName = ?, JobTitle = ?, Department = ?, City = ?, Country = ?,
                                                    TrustType = ?, TotalStorage = ?, FreeStorage = ?, PhysicalMemory = ?
                                                    WHERE DeviceName = ? AND SerialNumber = ?''', 
                                                (new_row['UserPrincipalName'], new_row['OperatingSystem'], new_row['EntraDeviceID'], new_row['IntuneDeviceID'],
                                                new_row['OSVersion'], new_row['ComplianceState'], source_value,
                                                new_row['Model'], new_row['Manufacturer'], new_row['MAC'], new_row['IntuneLastSync'], 
                                                report_time_str, new_row['Encryption'], new_row['UserDisplayName'], new_row['JobTitle'], new_row['Department'],                                                  
                                                new_row['City'], new_row['Country'], new_row['TrustType'], new_row['TotalStorage'], 
                                                new_row['FreeStorage'], new_row['PhysicalMemory'],
                                                device_name, serial_number))
                    else:
                        thread_cursor.execute('REPLACE INTO Devices VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', 
                                            (device_name, serial_number, new_row['UserPrincipalName'], new_row['EntraDeviceID'], new_row['IntuneDeviceID'],
                                            new_row['OperatingSystem'], new_row['OSVersion'], new_row['ComplianceState'], 'Cloud',
                                            new_row['Model'], new_row['Manufacturer'], new_row['MAC'], new_row['IntuneLastSync'],
                                            report_time_str, new_row['Encryption'], new_row['UserDisplayName'], new_row['JobTitle'], new_row['Department'], 
                                            new_row['City'], new_row['Country'], new_row['TrustType'], new_row['TotalStorage'], new_row['FreeStorage'], new_row['PhysicalMemory']))

                # Step: Identify records in the database but not in the latest Excel data and update Source to "Local"                
                thread_cursor.execute("SELECT SerialNumber, DeviceName, Source FROM Devices")
                all_db_records = thread_cursor.fetchall()

                for record in all_db_records:
                    serial_number, device_name, *rest = record
                    source = rest[0] if rest else None
                    if (serial_number.strip().lower(), device_name.strip().lower()) not in latest_records and source != 'Local':
                        thread_cursor.execute("UPDATE Devices SET Source = 'Local' WHERE SerialNumber = ? AND DeviceName = ?", (serial_number, device_name))

                thread_conn.commit()

            # Reload data from the database to reflect changes and update remarks accordingly
            self.load_data_for_refresh()
        except Exception as e:
            print(f"Failed to refresh data: {e}\n{traceback.format_exc()}")
        finally:
            # Stop and close the progress window
            progress.stop()
            progress_window.destroy()

    def remove_duplicates(self):
        # Clear any active search or filters before removing duplicates
        self.clear_filters()
        if self.data is not None:
            # Ensure ReportTime is in a consistent datetime format
            try:
                self.data['ReportTime'] = pd.to_datetime(self.data['ReportTime'], errors='coerce')
            except Exception as e:
                print(f"Failed to convert ReportTime to datetime: {e}")
                return

            # Group by all columns except ReportTime and identify duplicates
            non_datetime_columns = self.data.columns.difference(['ReportTime'])
            grouped = self.data.groupby(list(non_datetime_columns), sort=False)

            # Function to get the preferred record
            def get_preferred_record(group):
                if len(group) > 1:
                    # If there's a timestamp with 'T' and 'Z', remove it if another timestamp exists
                    filtered_group = group[~group['ReportTime'].astype(str).str.contains('T.*Z')] if len(group) > 1 else group
                    valid_dates = filtered_group.dropna(subset=['ReportTime'])
                    if not valid_dates.empty:
                        return valid_dates.loc[valid_dates['ReportTime'].idxmax()]
                return group.iloc[0]

            # Apply the function to remove duplicates
            resolved_data = grouped.apply(get_preferred_record).reset_index(drop=True)

            # Update the data and display it
            self.data = resolved_data
            self.display_data(self.data)
            self.update_total_records(self.data)
        else:
            print("No data to process.")

    def update_total_records(self, data):
        total_records = len(self.data)
        displayed_records = len(data)
        self.update_total_records_label(displayed_records, total_records)

    def update_total_records_label(self, displayed, total):
        # Update the displayed part with dynamic color
        if displayed == total:
            displayed_color = "#66CDAA"  # Soft green
        elif displayed < total:
            displayed_color = "#FFA07A"  # Soft orange
        else:
            displayed_color = "#66CDAA"  # Default to green if anything else occurs (which shouldn't happen)

        # Configure the label for the displayed count
        self.total_records_displayed_label.config(text=f"{displayed}", fg=displayed_color)
        
        # Configure the label for the total count (always green)
        self.total_records_total_label.config(text=f"/  {total}", fg="#66CDAA")

    def display_data(self, data):
        # Clear the existing treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Set up new treeview columns based on selected columns
        columns_present = [col for col in self.selected_columns if col in data.columns]
        self.tree["column"] = columns_present
        self.tree["show"] = "headings"

        for col in columns_present:
            self.tree.heading(col, text=col, command=lambda _col=col: self.sort_data(_col))
            header_width = max(100, len(col) * 10)  # Ensure a default minimum width for headers
            self.tree.column(col, width=header_width, minwidth=header_width, stretch=True, anchor='center')

        # Get only rows for the current page
        start_index = self.current_page * self.page_size
        end_index = start_index + self.page_size
        current_page_data = data.iloc[start_index:end_index]

        # Insert data into treeview
        for _, row in current_page_data.iterrows():
            row_values = row.copy()

            # For 'Local' records, convert Encryption value from 1/0 to True/False
            if row_values['Source'] == 'Local' and 'Encryption' in row_values:
                if row_values['Encryption'] == '1':
                    row_values['Encryption'] = True
                elif row_values['Encryption'] == '0':
                    row_values['Encryption'] = False

            self.tree.insert("", "end", values=[str(row_values[col]) for col in columns_present])

        # Adjust window width based on current columns
        self.adjust_window_width()

        # Update the display
        self.root.update()
        self.update_total_records(data)  # Update total records label after displaying data

    def next_page(self):
        total_pages = (len(self.filtered_data) - 1) // self.page_size + 1
        if (self.current_page + 1) < total_pages:
            self.current_page += 1
            self.display_data(self.filtered_data)
            self.update_page_number_label()

    def previous_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.display_data(self.filtered_data)
            self.update_page_number_label()

    def update_page_number_label(self):
        total_pages = (len(self.filtered_data) - 1) // self.page_size + 1
        current_page_display = self.current_page + 1  # Pages start at 1 for display purposes
        self.page_number_label.config(text=f"Page {current_page_display} / {total_pages}")

    def apply_filters(self):
        # Get the search query
        search_query = self.search_entry.get().strip().lower()

        # Start with the full dataset
        filtered_data = self.data.copy()

        # Apply the search query only to visible columns in the TreeView
        visible_columns = list(self.tree["columns"])
        if search_query and visible_columns:
            try:
                # Filter rows where the search query matches any value in visible columns
                filtered_data = filtered_data[
                    filtered_data[visible_columns].apply(
                        lambda row: search_query in ' '.join(row.astype(str).str.lower()), axis=1
                    )
                ]
            except KeyError as e:
                print(f"KeyError in search operation: {e}")
                messagebox.showerror("Error", "One or more displayed columns are missing in the data.")
                return

        # Initialize a list to store conditions for each filter
        conditions = []

        # Apply each individual filter
        for i in range(len(self.filter_column_comboboxes)):
            column = self.filter_column_comboboxes[i].get()  # Get the selected column for filtering
            operator = self.filter_operator_comboboxes[i].get()  # Get the selected operator
            value = self.filter_value_entries[i].get().strip()  # Get the filter value

            if column and operator and value:  # Apply filter only if all fields are filled
                try:
                    column_data = filtered_data[column].astype(str).str.lower()  # Normalize case
                except KeyError as e:
                    print(f"KeyError in filter operation: {e}")
                    messagebox.showerror("Error", f"Column '{column}' is missing in the data.")
                    return

                value = value.lower()

                # Apply the filter based on the operator
                if operator == "equals":
                    condition = column_data == value
                elif operator == "does not equal":
                    condition = column_data != value
                elif operator == "contains":
                    condition = column_data.str.contains(value, na=False)
                elif operator == "does not contain":
                    condition = ~column_data.str.contains(value, na=False)
                elif operator == "begins with":
                    condition = column_data.str.startswith(value, na=False)
                elif operator == "does not begin with":
                    condition = ~column_data.str.startswith(value, na=False)
                elif operator == "ends with":
                    condition = column_data.str.endswith(value, na=False)
                elif operator == "does not end with":
                    condition = ~column_data.str.endswith(value, na=False)
                else:
                    condition = None

                if condition is not None:
                    conditions.append(condition)

        # Combine conditions using the selected global operator
        if conditions:
            global_operator = self.operator_combobox.get().strip()
            if global_operator == "AND":
                # Combine all conditions with AND
                combined_condition = conditions[0]
                for condition in conditions[1:]:
                    combined_condition &= condition
            elif global_operator == "OR":
                # Combine all conditions with OR
                combined_condition = conditions[0]
                for condition in conditions[1:]:
                    combined_condition |= condition
            else:
                combined_condition = None

            if combined_condition is not None:
                filtered_data = filtered_data[combined_condition]

        # Update the filtered data and reset pagination
        self.filtered_data = filtered_data.copy()
        self.current_page = 0  # Reset to the first page
        self.display_data(self.filtered_data)  # Display the updated data
        self.update_total_records(self.filtered_data)  # Update record count
        self.update_page_number_label()  # Update pagination display

    def clear_filters(self):
        if self.data is not None:
            self.filtered_data = self.data.copy()
            self.current_page = 0  # Reset to the first page
            self.display_data(self.filtered_data)
            self.update_total_records(self.filtered_data)
            self.update_page_number_label()  # Update the page number label

            # Clear search and filter inputs
            self.search_entry.delete(0, tk.END)
            for i in range(len(self.filter_column_comboboxes)):
                self.filter_column_comboboxes[i].set("")
                self.filter_operator_comboboxes[i].set("equals")  # Reset to default operator
                self.filter_value_entries[i].delete(0, tk.END)  # Clear the text entry
            self.operator_combobox.set("AND")
        else:
            messagebox.showinfo("Info", "Please load data first.")

    def sort_data(self, column):
        if self.filtered_data is not None:
            if column:
                try:
                    # Determine sort order (toggle between ascending and descending)
                    if column in self.sort_orders and self.sort_orders[column] == 'asc':
                        ascending = False
                        self.sort_orders[column] = 'desc'
                    else:
                        ascending = True
                        self.sort_orders[column] = 'asc'

                    # Convert the column to string to avoid mixed type comparison issues
                    self.filtered_data[column] = self.filtered_data[column].astype(str)

                    # Sort the entire filtered dataset
                    sorted_data = self.filtered_data.sort_values(by=[column], ascending=ascending)

                    # Update the filtered data and reset pagination
                    self.filtered_data = sorted_data.copy()
                    self.current_page = 0  # Reset to the first page after sorting
                    
                    # Display data with the currently selected columns maintained
                    self.display_data(self.filtered_data)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to sort data: {e}")
    
    def sort_software_data(self, column):
        if self.software_filtered_data is not None:
            if column:
                try:
                    # Determine sort order (toggle between ascending and descending)
                    if column in self.software_sort_orders and self.software_sort_orders[column] == 'asc':
                        ascending = False
                        self.software_sort_orders[column] = 'desc'
                    else:
                        ascending = True
                        self.software_sort_orders[column] = 'asc'

                    # Convert the column to string to avoid mixed type comparison issues
                    self.software_filtered_data[column] = self.software_filtered_data[column].astype(str)

                    # Sort the entire filtered dataset
                    sorted_data = self.software_filtered_data.sort_values(by=[column], ascending=ascending)

                    # Update the filtered data and reset pagination
                    self.software_filtered_data = sorted_data.copy()
                    self.software_current_page = 0  # Reset to the first page after sorting
                    
                    # Display data with the currently selected columns maintained
                    self.display_software_data(self.software_filtered_data)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to sort data: {e}")

    def show_context_menu(self, event):
        # Get the selected item
        item_id = self.tree.identify_row(event.y)
        if item_id:
            self.tree.selection_set(item_id)
            self.context_menu.post(event.x_root, event.y_root)

    def show_software_context_menu(self, event):
        # Get the selected item
        item_id = self.software_tree.identify_row(event.y)
        if item_id:
            self.software_tree.selection_set(item_id)
            self.software_context_menu.post(event.x_root, event.y_root)

    def copy_cell_value(self):
        selected_item = self.tree.selection()
        if selected_item:
            item_values = self.tree.item(selected_item, "values")
            column_index = self.tree.identify_column(self.tree.winfo_pointerx() - self.tree.winfo_rootx()).replace("#", "")
            if column_index.isdigit():
                column_index = int(column_index) - 1
                if 0 <= column_index < len(item_values):
                    value = item_values[column_index]
                    self.root.clipboard_clear()
                    self.root.clipboard_append(value)
                    self.root.update()  # Now it stays on the clipboard

    def copy_software_cell_value(self):
        selected_item = self.software_tree.selection()
        if selected_item:
            item_values = self.software_tree.item(selected_item, "values")
            column_index = self.software_tree.identify_column(self.software_tree.winfo_pointerx() - self.software_tree.winfo_rootx()).replace("#", "")
            if column_index.isdigit():
                column_index = int(column_index) - 1
                if 0 <= column_index < len(item_values):
                    value = item_values[column_index]
                    self.root.clipboard_clear()
                    self.root.clipboard_append(value)
                    self.root.update()  # Now it stays on the clipboard

    def add_edit_remark(self):
        selected_item = self.tree.selection()
        if selected_item:
            item_values = self.tree.item(selected_item, "values")
            # Create a mapping of column names to their index positions
            columns_mapping = {col: idx for idx, col in enumerate(self.tree["columns"])}
            serial_number = item_values[columns_mapping["SerialNumber"]].strip().lower()
            device_name = item_values[columns_mapping["DeviceName"]].strip().lower()

            # Create a new window for adding/editing remark
            remark_window = tk.Toplevel(self.root)
            remark_window.title("Add/Edit Remark")
            remark_window.geometry("400x200")
            remark_window.transient(self.root)
            remark_window.grab_set()  # Make the main window unresponsive

            # Center the remark window on the screen
            remark_window.update_idletasks()
            x = (remark_window.winfo_screenwidth() // 2) - (remark_window.winfo_width() // 2)
            y = (remark_window.winfo_screenheight() // 2) - (remark_window.winfo_height() // 2)
            remark_window.geometry(f"+{x}+{y}")

            tk.Label(remark_window, text="Enter Remark:", font=("Arial", 12)).pack(pady=10)
            remark_entry = tk.Entry(remark_window, font=("Arial", 12), width=40)
            remark_entry.pack(pady=5)
            remark_entry.focus_set()  # Set focus to the entry box

            # Pre-fill the existing remark if available
            self.cursor.execute("SELECT Remarks FROM Remarks WHERE SerialNumber = ? AND DeviceName = ?", (serial_number, device_name))
            existing_remark = self.cursor.fetchone()
            if existing_remark:
                remark_entry.insert(0, existing_remark[0])

            def save_remark():
                new_remark = remark_entry.get()
                if new_remark:
                    # Update or insert the remark in the database using the correct Serial Number and Device Name
                    self.cursor.execute("REPLACE INTO Remarks (SerialNumber, DeviceName, Remarks) VALUES (?, ?, ?)",
                                        (serial_number, device_name, new_remark))
                    self.conn.commit()

                    # Update the remark in both self.data and self.filtered_data DataFrames to keep them consistent
                    self.data.loc[(self.data['SerialNumber'].str.strip().str.lower() == serial_number) &
                                (self.data['DeviceName'].str.strip().str.lower() == device_name), 'Remarks'] = new_remark

                    # Update the database with the modified Source value
                    self.conn.commit()

                    # Update filtered data directly if the item is present
                    if not self.filtered_data.empty:
                        self.filtered_data.loc[(self.filtered_data['SerialNumber'].str.strip().str.lower() == serial_number) &
                                            (self.filtered_data['DeviceName'].str.strip().str.lower() == device_name), 'Remarks'] = new_remark

                    # Refresh the displayed data to reflect the changes
                    self.display_data(self.filtered_data)
                    self.update_total_records(self.filtered_data)

                remark_window.destroy()

            save_button = tk.Button(remark_window, text="Save", command=save_remark, font=("Arial", 12), bg="#4A5568", fg="#E4E7EB", activebackground="#6B7280")
            save_button.pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = DeviceInventoryApp(root)
    root.mainloop()
