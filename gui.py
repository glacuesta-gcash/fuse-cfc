import tkinter as tk
from tkinter import filedialog, messagebox
import sys

def run_commands():
    credentials_path = credentials_var.get()
    hex_id = hex_id_var.get()
    
    # Validate that the file and hex ID are provided
    if not credentials_path:
        messagebox.showerror("Error", "Please select a credentials file.")
        return
    if not hex_id:
        messagebox.showerror("Error", "Please enter a hex ID.")
        return
    
    # Run your set of commands here
    # Replace this with your own function
    print(f"Running commands with credentials file: {credentials_path} and hex ID: {hex_id}")
    messagebox.showinfo("Info", "Commands executed successfully!")

def select_file():
    file_path = filedialog.askopenfilename(title="Select Credentials File")
    credentials_var.set(file_path)

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

# Set the icon
icon_path = "icon-24.png"
root.iconphoto(True, tk.PhotoImage(file=icon_path))

# Variables to hold the inputs
credentials_var = tk.StringVar()
hex_id_var = tk.StringVar()

# Credentials file field
tk.Label(root, text="Credentials File:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
tk.Entry(root, textvariable=credentials_var, width=50).grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="Browse", command=select_file).grid(row=0, column=2, padx=5, pady=5)

# Hex ID field
tk.Label(root, text="Google Sheet ID:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
tk.Entry(root, textvariable=hex_id_var, width=50).grid(row=1, column=1, padx=5, pady=5)

# Run button
tk.Button(root, text="Run", command=run_commands).grid(row=2, column=1, pady=10)

# Output text box
tk.Label(root, text="Output:").grid(row=3, column=0, sticky="ne", padx=5, pady=(5,20))
output_text = tk.Text(root, height=10, width=40, bg=root.cget("bg"), state="disabled")
output_text.grid(row=3, column=1, columnspan=2, padx=5, pady=5, sticky="w")

# Add a label with plain text
plain_text_label = tk.Label(root, text="For questions please contact gabby.lacuesta@gcash.com")
plain_text_label.grid(row=4, columnspan=3, padx=10, pady=10)

# Redirect print output to the Text widget
sys.stdout = TextRedirector(output_text)

# Start the Tkinter main loop
root.mainloop()