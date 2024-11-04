import tkinter as tk
from tkinter import filedialog, messagebox, font
import sys
import os
import re

from sheet import initialize_sheets, Sheet
from commands import Command
from timer import Timer

def run_commands():
    credentials_path = credentials_var.get()
    g_url = gsheet_url_var.get()
    
    # Validate that the file and hex ID are provided
    if not credentials_path:
        messagebox.showerror("Error", "Please select a credentials file.")
        return
    if not g_url:
        messagebox.showerror("Error", "Please enter a Google Sheet URL.")
        return
    
    match = re.search(r"/d/([a-zA-Z0-9_\-]+)", g_url)
    if match:
        g_id = match.group(1)
    else:
        messagebox.showerror("Error", "Invalid Google Sheet URL.")
        return

    # Run your set of commands here
    # Replace this with your own function
    print(f"Running commands...\n  Credentials: {credentials_path}\n  Google Sheet ID: {g_id}\n\n")

    timer = Timer()
    initialize_sheets(credentials_path)
    sheet = Sheet(g_id)

    cmd = sheet.steps_tab.read_next_command()
    while cmd is not None:
        Command(cmd).exec(sheet)
        cmd = sheet.steps_tab.read_next_command()

    sheet.summarize()

    sheet.flush()

    print(f'\n✔ Done {timer.check()}')

    messagebox.showinfo("Info", "Commands executed successfully!")

def select_file():
    file_path = filedialog.askopenfilename(title="Select Credentials File")
    credentials_var.set(file_path)

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def get_system_font():
    # Check if ".SF NS Text" is available (macOS system font)
    if ".SF NS Text" in font.families():
        return (".SF NS Text", 10)
    # If not, fallback to "Segoe UI" (Windows system font)
    elif "Segoe UI" in font.families():
        return ("Segoe UI", 10)
    # Default fallback to a common font
    else:
        return ("Arial", 10)
    
class TextRedirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        self.widget.config(state="normal")  # Enable editing
        self.widget.insert(tk.END, text)
        self.widget.see(tk.END)  # Auto-scroll to the end
        self.widget.config(state="disabled")  # Make read-only

    def flush(self):
        pass  # Needed for file-like behavior but can be left empty

# Create the main window
root = tk.Tk()
root.title("Cascading Forecasts")

root.update()

# Set the icon
# for mac, use icon-256; windows, use icon-24
icon_path = resource_path("icon-256.png")
root.iconphoto(True, tk.PhotoImage(file=icon_path))

# Variables to hold the inputs
credentials_var = tk.StringVar()
gsheet_url_var = tk.StringVar()

# Credentials file field
tk.Label(root, text="Credentials:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
tk.Entry(root, textvariable=credentials_var, width=50).grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="Browse", command=select_file).grid(row=0, column=2, padx=5, pady=5)

# Hex ID field
tk.Label(root, text="Google Sheet URL:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
tk.Entry(root, textvariable=gsheet_url_var, width=50).grid(row=1, column=1, padx=5, pady=5)

# Run button
tk.Button(root, text="Run", command=run_commands).grid(row=2, column=1, pady=10)

# # Output text box
# tk.Label(root, text="Output:").grid(row=3, column=0, sticky="nw", padx=5, pady=5)
# output_text = tk.Text(root, height=10, width=64, bg=root.cget("bg"), state="disabled", font=get_system_font())
# output_text.grid(row=4, column=0, columnspan=3, padx=5, pady=5)

# Redirect print output to the Text widget
# TODO: removed this as it causes the GUI to hang
#sys.stdout = TextRedirector(output_text)

# Add a label with plain text
plain_text_label = tk.Label(root, text="For questions please contact gabby.lacuesta@gcash.com")
plain_text_label.grid(row=5, columnspan=3, padx=10, pady=10)

# Start the Tkinter main loop
root.mainloop()

# To compile via PyInstaller:
# win: pyinstaller --onefile --windowed --icon=favicon-256.ico --add-data "icon-24.png;." gui.py
# mac: pyinstaller --onefile --windowed --icon=icon-256.icns --add-data "icon-256.png:icon-24.png" gui.py