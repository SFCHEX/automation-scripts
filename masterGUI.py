import tkinter as tk
from tkinter import ttk, filedialog
import subprocess

script_names = ["outageMaster.py", "availabilityMaster.py", "powerMaster.py", "AVAPADPARMaster.py"]

def select_directory():
    directory_path = filedialog.askdirectory()
    directory_var.set(directory_path)

def run_script(script_name):
    directory_path = directory_var.get()
    if directory_path:
        try:
            result = subprocess.run(["python", script_name], cwd=directory_path, text=True, capture_output=True)
            output_text.config(state="normal")
            output_text.delete("1.0", tk.END)
            output_text.insert(tk.END, result.stdout)
            output_text.insert(tk.END, result.stderr)
            output_text.config(state="disabled")
        except FileNotFoundError:
            print(f"Script {script_name} not found.")
    else:
        print("Please select a directory first.")

root = tk.Tk()
root.title("Script Runner")

# Disable resizing
root.resizable(False, False)

# Use a more modern theme
style = ttk.Style()

# Configure padding and font
style.configure("TButton", padding=10, relief="flat", background="#007acc", foreground="black", font=("Helvetica", 12))
style.configure("TLabel", padding=6, font=("Helvetica", 12))

directory_label = ttk.Label(root, text="Select Directory:")
directory_label.grid(row=0, column=0, padx=10, pady=10)
select_button = ttk.Button(root, text="Browse", command=select_directory)
select_button.grid(row=0, column=1, padx=10, pady=10)

row = 1
col = 0
for script_name in script_names:
    button = ttk.Button(root, text=f"Run Script {script_name.upper()}", command=lambda s=script_name: run_script(s))
    button.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
    col += 1
    if col > 1:
        col = 0
        row += 1

# Create a resizable frame for the output text
output_frame = ttk.Frame(root)
output_frame.grid(row=row + 1, columnspan=2, padx=10, pady=10, sticky="nsew")
output_frame.columnconfigure(0, weight=1)
output_frame.rowconfigure(0, weight=1)

directory_var = tk.StringVar()

root.mainloop()
