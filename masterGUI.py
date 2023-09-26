import tkinter as tk
from tkinter import ttk, filedialog
import subprocess
import threading

script_names = ["outageMaster.py", "availabilityMaster.py", "powerMaster.py", "AVAPADPARMaster.py"]

script_dir = "C:\\Users\\swx1283483\\automation-scripts\\"

def select_directory():
    directory_path = filedialog.askdirectory()
    directory_var.set(directory_path)
    
def run_script(script_name):
    directory_path = directory_var.get()
    if directory_path:
        try:
            def run_subprocess():
                process = subprocess.Popen(["python", script_dir + script_name], cwd=directory_path, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                
                # Capture and print the output of the subprocess
                for line in process.stdout:
                    print(line, end='')
                for line in process.stderr:
                    print(line, end='')

            threading.Thread(target=run_subprocess).start()
        except FileNotFoundError:
            print(f"Script {script_name} not found.")
    else:
        print("Please select a directory first.")


root = tk.Tk()
root.title(directory_var)

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
        
directory_var = tk.StringVar()


root.mainloop()
