import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import subprocess
import webbrowser
import threading

# Define paths
base_folder = os.path.dirname(os.path.abspath(__file__))
input_folder = os.path.join(base_folder, "InputFolder")
output_folder = os.path.join(base_folder, "OutputFolder")

# Ensure folders exist
os.makedirs(input_folder, exist_ok=True)
os.makedirs(output_folder, exist_ok=True)

def update_file_lists():
    """Update the input and output file lists."""
    input_listbox.delete(0, tk.END)
    output_listbox.delete(0, tk.END)

    # List files in InputFolder
    for file in os.listdir(input_folder):
        input_listbox.insert(tk.END, file)

    # List files in OutputFolder
    for file in os.listdir(output_folder):
        output_listbox.insert(tk.END, file)

def upload_files():
    """Open file dialog to select files and copy them to InputFolder."""
    files = filedialog.askopenfilenames(title="Select Files to Upload", filetypes=[("Excel Files", "*.xlsx")])
    if files:
        for file in files:
            try:
                shutil.copy(file, input_folder)  # Copy files to InputFolder
            except Exception as e:
                messagebox.showerror("Error", f"Failed to copy {file}: {e}")
        messagebox.showinfo("Success", f"{len(files)} file(s) uploaded to InputFolder.")
        update_file_lists()

def run_program():
    """Run the main program and update the output file list."""
    def process():
        try:
            # Start the progress bar
            progress_bar.start()

            # Run the main program (CleanDataKontoutskrift.py)
            subprocess.run(["python", "CleanDataKontoutskrift.py"], check=True)
            messagebox.showinfo("Success", "Program executed successfully!")
            update_file_lists()
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"Program execution failed: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        finally:
            # Stop the progress bar
            progress_bar.stop()

    # Run the process in a separate thread to avoid freezing the GUI
    threading.Thread(target=process).start()

def open_file_from_listbox(listbox, folder):
    """Open the selected file from the specified folder."""
    try:
        selected_file = listbox.get(listbox.curselection())
        file_path = os.path.join(folder, selected_file)
        if os.path.exists(file_path):
            webbrowser.open(file_path)  # Open the file with the default application
        else:
            messagebox.showerror("Error", f"File not found: {file_path}")
    except tk.TclError:
        messagebox.showerror("Error", "No file selected.")

# Create the main application window
root = tk.Tk()
root.title("CleanupAtlas GUI")
root.geometry("700x700")  # Set the window size

# Add a label
label = tk.Label(root, text="Welcome to CleanupAtlas!", font=("Times", 24))
label.pack(pady=10)

# Create a frame for the input and output lists
list_frame = tk.Frame(root)
list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

# Input file list
input_label = tk.Label(list_frame, text="Input Files:", font=("Arial", 12))
input_label.grid(row=0, column=0, sticky="w")
input_listbox = tk.Listbox(list_frame, width=50, height=15)
input_listbox.grid(row=1, column=0, padx=5, pady=5)
input_listbox.bind("<Double-Button-1>", lambda event: open_file_from_listbox(input_listbox, input_folder))

# Output file list
output_label = tk.Label(list_frame, text="Output Files:", font=("Arial", 12))
output_label.grid(row=2, column=0, sticky="w")
output_listbox = tk.Listbox(list_frame, width=50, height=15)
output_listbox.grid(row=3, column=0, padx=5, pady=5)
output_listbox.bind("<Double-Button-1>", lambda event: open_file_from_listbox(output_listbox, output_folder))

# Create a frame for the buttons
button_frame = tk.Frame(list_frame)
button_frame.grid(row=1, column=1, rowspan=3, padx=10, pady=5, sticky="n")

# Add buttons to the button frame
upload_button = tk.Button(button_frame, text="Upload Files", command=upload_files, width=15, height=2)
upload_button.pack(pady=10)

refresh_button = tk.Button(button_frame, text="Refresh", command=update_file_lists, width=15, height=2)
refresh_button.pack(pady=10)

run_button = tk.Button(button_frame, text="Run", command=run_program, width=15, height=2)
run_button.pack(pady=20)

# Add a progress bar
progress_bar = ttk.Progressbar(root, mode="indeterminate", length=400)
progress_bar.pack(pady=20)

# Initialize file lists
update_file_lists()

# Run the application
root.mainloop()