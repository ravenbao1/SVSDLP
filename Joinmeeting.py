import pygetwindow as gw
import pyautogui
import time
import sys
import argparse
import screeninfo
import os

# Path to the log file
log_file_path = r"C:\temp\meetings\log.txt"

# Desired window size
desired_width = 1000
desired_height = 650

def log_message(message):
    """Log messages to the file."""
    with open(log_file_path, "a") as log_file:
        log_file.write(message + "\n")

def focus_window(window_title):
    """ Bring the window to the foreground and resize it based on the title """
    all_windows = gw.getAllTitles()
    
    # Find the target window
    target_window = next((window for window in all_windows if window_title in window), None)

    if not target_window:
        log_message(f"Window with title containing '{window_title}' not found.")
        return None

    # Activate the target window
    window = gw.getWindowsWithTitle(target_window)[0]
    window.restore()  # Ensure the window is not minimized
    window.activate()
    time.sleep(1)  # Wait for the window to come to the foreground

    # Resize the window to the desired size (1000x650)
    log_message(f"Resizing window to {desired_width}x{desired_height}")
    window.resizeTo(desired_width, desired_height)
    time.sleep(2)
    return window

def get_display_coordinates():
    """Retrieve the coordinates of all displays."""
    monitors = screeninfo.get_monitors()
    for i, monitor in enumerate(monitors):
        log_message(f"Monitor {i}: {monitor}")
    return monitors

def get_monitor_for_window(window):
    """ Find which monitor the window is primarily located on """
    monitors = get_display_coordinates()
    
    for monitor in monitors:
        # Check if the window is within the bounds of the monitor
        if monitor.x <= window.left <= monitor.x + monitor.width and \
           monitor.y <= window.top <= monitor.y + monitor.height:
            log_message(f"Window found on monitor: {monitor}")
            return monitor
    
    log_message("Window not found on any monitor.")
    return None

def click_at_coordinates(window, click_x, click_y):
    """ Click at the specified coordinates within the window """
    try:
        # Get window position on the screen
        window_left = window.left
        window_top = window.top

        # Calculate the absolute coordinates by adding the window's top-left position
        absolute_x = window_left + click_x
        absolute_y = window_top + click_y

        log_message(f"Window position: ({window_left}, {window_top})")
        log_message(f"Window size: ({window.width}, {window.height})")

        # Find which monitor the window is on
        monitor = get_monitor_for_window(window)
        
        if monitor:
            # Check if the calculated coordinates are within the monitor bounds
            if monitor.x <= absolute_x <= monitor.x + monitor.width and \
               monitor.y <= absolute_y <= monitor.y + monitor.height:
                log_message(f"Clicking at window-relative coordinates: ({click_x}, {click_y})")
                log_message(f"Clicking at absolute screen coordinates: ({absolute_x}, {absolute_y})")
                pyautogui.click(absolute_x, absolute_y)
                log_message("Click successfully performed.")
            else:
                log_message("Calculated click coordinates are outside of the monitor bounds.")
                click_at_window_center(window)
        else:
            log_message("Unable to determine which monitor the window is on. Clicking at the window center.")
            click_at_window_center(window)

    except Exception as e:
        log_message(f"Unexpected error occurred: {e}")
        raise

def click_at_window_center(window):
    """ Fallback: Click at the center of the window """
    try:
        # Get the window center position
        window_center_x = window.left + window.width // 2
        window_center_y = window.top + window.height // 2

        log_message(f"Clicking at window center coordinates: ({window_center_x}, {window_center_y})")
        
        # Perform the click at the window center
        pyautogui.click(window_center_x, window_center_y)
        log_message("Click successfully performed.")
    except Exception as e:
        log_message(f"Error while trying to click at window center: {e}")
        raise

if __name__ == "__main__":
    # Ensure the log directory exists
    os.makedirs(r"C:\temp\meetings", exist_ok=True)

    # Set up argument parsing
    parser = argparse.ArgumentParser(description="Script to focus on a window, resize it, and click specified coordinates.")
    parser.add_argument("window_title", type=str, help="The title of the window to focus on (e.g., 'test44').")
    parser.add_argument("click_x", type=int, help="X coordinate to click within the window (relative).")
    parser.add_argument("click_y", type=int, help="Y coordinate to click within the window (relative).")

    # Parse the arguments
    args = parser.parse_args()

    # Get the window title and coordinates from the passed arguments
    window_title = f"{args.window_title} | Microsoft Teams"
    click_x = args.click_x
    click_y = args.click_y

    # Focus on the specific window and resize it
    focused_window = focus_window(window_title)

    if focused_window:
        log_message(f"Window '{window_title}' is now in focus and resized.")
        # Click at the specified coordinates within the window
        click_at_coordinates(focused_window, click_x, click_y)
    else:
        log_message("Window not found. Exiting.")
