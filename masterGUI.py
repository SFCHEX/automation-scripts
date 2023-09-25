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
                while True:
                    output = process.stdout.readline()
                    if output == '' and process.poll() is not None:
                        break
                    if output:
                        print(output.strip())  # Print to console
                        output_queue.put(output)  # Put output into the queue for UI update

            threading.Thread(target=run_subprocess).start()
        except FileNotFoundError:
            print(f"Script {script_name} not found.")
    else:
        print("Please select a directory first.")

def update_output():
    while True:
        output = output_queue.get()
        if output is None:
            break
        output_text_widget.config(state="normal")
        output_text_widget.insert(tk.END, output)
        output_text_widget.config(state="disabled")
        output_text_widget.see(tk.END)  # Scroll to the end
        root.update_idletasks()  # Update the UI

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

output_frame = ttk.Frame(root)
output_frame.grid(row=row + 1, columnspan=2, padx=10, pady=10, sticky="nsew")
output_frame.columnconfigure(0, weight=1)
output_frame.rowconfigure(0, weight=0)  # Set the weight to 0 to make it smaller

directory_var = tk.StringVar()

output_text_widget = tk.Text(output_frame, wrap=tk.WORD, state="disabled", font=("Helvetica", 12))
output_text_widget.grid(row=0, column=0, sticky="nsew")

output_queue = queue.Queue()  # Create a queue for output

# Start the output update thread
output_thread = threading.Thread(target=update_output)
output_thread.daemon = True  # Allow the thread to exit when the main program exits
output_thread.start()

root.mainloop()
