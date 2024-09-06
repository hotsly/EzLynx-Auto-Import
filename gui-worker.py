import customtkinter as ctk
import os
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

# Global variables
file_path = ""
statement_number = ""  # Global variable for statement number
driver = None  # Global variable to hold the WebDriver instance

def create_user_data_directory():
    """Create the User Data directory if it does not exist."""
    user_data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'User Data')
    if not os.path.exists(user_data_dir):
        os.makedirs(user_data_dir)

def wait_for_login(driver, wait):
    login_button_selector = By.ID, 'btnLogin'
    target_url = 'https://app.ezlynx.com/ApplicantPortal/Commissions/CommissionStatement/ImportStatement'

    while True:
        current_url = driver.current_url
        if current_url == target_url:
            print('Target URL reached. Refreshing the page twice.')
            driver.refresh()
            break

        try:
            # Check if login button is present
            wait.until(EC.presence_of_element_located(login_button_selector))
            print('Login button found. Please log in.')
            time.sleep(5)
        except Exception:
            print('Login button not found or another issue.')
            time.sleep(5)

def format_date(date_str):
    parts = date_str.split('/')
    month = str(int(parts[0]))
    day = str(int(parts[1]))
    year = parts[2].zfill(4)
    return f"{month}/{day}/{year}"

def get_current_date():
    """Get today's date formatted as MM/DD/YYYY."""
    today = datetime.now()
    return today.strftime("%m/%d/%Y")

def process_file():
    global driver
    global statement_number

    if not file_path:
        messagebox.showerror("Error", "No file selected.")
        return

    if not statement_number:
        messagebox.showwarning("Warning", "Statement number is not provided.")
        return

    # Disable buttons and update status
    browse_button.configure(state=tk.DISABLED)
    start_button.configure(state=tk.DISABLED)
    status_label.configure(text="Processing...")
    root.update_idletasks()

    try:
        # Create user data directory
        create_user_data_directory()

        # Define paths
        script_dir = os.path.dirname(os.path.abspath(__file__))
        chrome_driver_dir = os.path.join(script_dir, 'chromedriver-win64')
        chrome_driver_path = os.path.join(chrome_driver_dir, 'chromedriver.exe')
        user_data_dir = os.path.join(script_dir, 'User Data')

        # Setup Chrome options
        chrome_options = Options()
        chrome_options.add_argument(f"user-data-dir={user_data_dir}")
        chrome_options.add_argument('profile-directory=Default')
        chrome_options.add_argument("--window-size=1024,768")

        # Initialize the WebDriver
        driver = webdriver.Chrome(service=Service(chrome_driver_path), options=chrome_options)
        wait = WebDriverWait(driver, 10)

        driver.get("https://app.ezlynx.com/ApplicantPortal/Commissions/CommissionStatement/ImportStatement")

        # Use the wait_for_login function to ensure the user is logged in
        wait_for_login(driver, wait)
        
        driver.get("https://app.ezlynx.com/ApplicantPortal/Commissions/CommissionStatement/ImportStatement")

        # Click on 'Upload File' button
        upload_button = driver.find_element(By.XPATH, '//button[@ng-click="model.NavigateToTab(model.currentTabIndex + 1, true)"]')
        upload_button.click()

        # Load the CSV file
        df = pd.read_csv(file_path)

        # Extract the sums for columns E and F
        try:
            premium_sum = df['Premium Paid'].sum()
            commission_sum = df['Producer Split'].sum()
        except KeyError as e:
            messagebox.showerror("Error", f"Column not found: {e}")
            return
        except IndexError as e:
            messagebox.showerror("Error", f"Index error: {e}")
            return

        # Fill in the form
        statement_number_input = driver.find_element(By.XPATH, '//input[@ng-model="model.SummaryStatementNumber"]')
        statement_number_input.send_keys(statement_number)

        date_input = driver.find_element(By.XPATH, '//input[@id="SummaryStatementDate"]')
        current_date = get_current_date()
        formatted_date = format_date(current_date)
        date_input.send_keys(formatted_date)

        premium_input = driver.find_element(By.XPATH, '//input[@id="Premium"]')
        premium_input.send_keys(str(premium_sum))

        commission_input = driver.find_element(By.XPATH, '//input[@id="Commission"]')
        commission_input.send_keys(str(commission_sum))

        comments_input = driver.find_element(By.XPATH, '//textarea[@ng-model="model.SummaryComments"]')
        comments_input.send_keys(os.path.basename(file_path).replace("Approved", "").replace(".csv", "").strip())

        # Confirm the details
        if not messagebox.askokcancel("Confirm", "Is all the information correct?"):
            print("User cancelled the operation.")
            driver.quit()  # Close the WebDriver immediately
            status_label.configure(text="Done!")
            browse_button.configure(state=tk.NORMAL)
            start_button.configure(state=tk.NORMAL)
            return

        upload_button.click()

        # Check for the Finish button and wait for user to press it
        finish_button_selector = By.XPATH, '//button[@ng-click="model.NavigateToTab(model.currentTabIndex + 1, true)"]'

        while True:
            try:
                finish_button = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(finish_button_selector)
                )
                if finish_button.is_displayed():
                    print("Finish button found. Please press it on the web page.")
                    break
            except Exception as e:
                print(f"Finish button not found or another issue: {e}")

            # Check if the WebDriver instance is still alive
            try:
                driver.current_window_handle  # This will raise an exception if the browser is closed
            except Exception as e:
                print(f"WebDriver is no longer available: {e}")
                messagebox.showwarning("Warning", "The WebDriver has been closed. Please restart the process.")
                break

    finally:
        if driver:
            driver.quit()
        # Enable buttons and update status
        browse_button.configure(state=tk.NORMAL)
        start_button.configure(state=tk.NORMAL)
        status_label.configure(text="Done!")


def start_processing_thread():
    # Retrieve the latest statement number value
    global statement_number
    statement_number = statement_number_entry.get().strip()

    if not file_path:
        messagebox.showerror("Error", "No file selected.")
        return

    if not statement_number:
        messagebox.showwarning("Warning", "Statement number is not provided.")
        return

    # Start the file processing in a new thread
    threading.Thread(target=process_file, daemon=True).start()

def browse_file():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        # Extract the file name and update the entry
        file_name = os.path.basename(file_path)
        file_path_entry.configure(state=tk.NORMAL)  # Temporarily make entry editable
        file_path_entry.delete(0, tk.END)
        file_path_entry.insert(0, file_name)
        file_path_entry.configure(state=tk.DISABLED)  # Make entry read-only
        start_button.configure(state=tk.NORMAL)  # Enable start button

def update_statement_number(event=None):
    global statement_number
    statement_number = statement_number_entry.get().strip()

# Create the GUI
root = ctk.CTk()  # Use CustomTkinter's CTk window
root.title("CSV Upload and Process")
root.geometry("500x300")  # Extended length for a more spacious look

# Setup customtkinter theme
ctk.set_appearance_mode("dark")  # Set dark mode
ctk.set_default_color_theme("dark-blue")  # Dark blue theme

# Main window background color
root.configure(bg="#2e2e2e")  # Dark gray background for the main window

# Create a compact frame to contain the status label
status_frame = ctk.CTkFrame(root, corner_radius=5, height=30, width=120)  # Adjusted to fit the label size
status_frame.place(relx=1.0, rely=1.0, anchor="se", x=0, y=0)  # Positioned at bottom-right corner

# Status Label inside the frame
status_label = ctk.CTkLabel(status_frame, text="Waiting for file...", text_color="white", font=("Arial", 10))
status_label.pack(pady=5, padx=5)

# File Name Label
file_name_label = ctk.CTkLabel(root, text="File Name:", text_color="white")
file_name_label.pack(pady=(10, 0), padx=20, anchor="w")

# File Path Entry
file_path_entry = ctk.CTkEntry(root, placeholder_text="No file selected", width=300, height=30, state=tk.DISABLED)
file_path_entry.pack(pady=(0, 10), padx=20)

# Statement Number Label
statement_number_label = ctk.CTkLabel(root, text="Statement Number:", text_color="white")
statement_number_label.pack(pady=(10, 0), padx=20, anchor="w")

# Statement Number Entry
statement_number_entry = ctk.CTkEntry(root, placeholder_text="Enter statement number", width=300, height=30)
statement_number_entry.pack(pady=(0, 10), padx=20)
statement_number_entry.bind("<KeyRelease>", update_statement_number)  # Update statement number on key release

# Browse Button
button_width = 140
button_height = 35
browse_button = ctk.CTkButton(root, text="Browse", command=browse_file, width=button_width, height=button_height)
browse_button.pack(pady=(10, 5))

# Start Button
start_button = ctk.CTkButton(root, text="Start", command=start_processing_thread, state=tk.DISABLED, width=button_width, height=button_height)
start_button.pack(pady=10)

# Start the Tkinter event loop
root.mainloop()
