import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import subprocess
import webbrowser
import threading
import pandas as pd

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
    output_files = os.listdir(output_folder)
    for file in output_files:
        output_listbox.insert(tk.END, file)

    # Enable or disable the Budget Creator button based on output files
    if output_files:
        budget_button.config(state=tk.NORMAL)
    else:
        budget_button.config(state=tk.DISABLED)

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

def open_budget_creator():
    """Open the Budget Creator window."""
    # Create a new window
    budget_window = tk.Toplevel(root)
    budget_window.title("Budget Creator")
    budget_window.geometry("800x800")

    # Add a label
    label = tk.Label(budget_window, text="Budget Creator", font=("Arial", 18))
    label.pack(pady=10)

    # Add a frame for the budget table
    table_frame = tk.Frame(budget_window)
    table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # Create a Treeview widget to display the budget data
    tree = ttk.Treeview(table_frame, columns=("Category", "Ut fra konto", "Inn på konto"), show="headings")
    tree.heading("Category", text="Category")
    tree.heading("Ut fra konto", text="Ut fra konto")
    tree.heading("Inn på konto", text="Inn på konto")
    tree.column("Category", width=200)
    tree.column("Ut fra konto", width=150)
    tree.column("Inn på konto", width=150)
    tree.pack(fill=tk.BOTH, expand=True)

    # Load data from Totals.xlsx
    try:
        totals_file = os.path.join(output_folder, "Totals.xlsx")
        if not os.path.exists(totals_file):
            raise FileNotFoundError("Totals.xlsx not found. Please run the program first.")

        # Read the Excel file
        df = pd.read_excel(totals_file, engine="openpyxl")

        # Insert data into the Treeview
        for _, row in df.iterrows():
            tree.insert("", tk.END, values=(row["Category"], row["Ut fra konto"], row["Inn på konto"]))

    except Exception as e:
        messagebox.showerror("Error", f"Failed to load budget data: {e}")

    # Add a header question
    question_label = tk.Label(budget_window, text="How much do you want to spend on:", font=("Arial", 14))
    question_label.pack(pady=10)

    # Create a scrollable frame for the input fields
    scrollable_frame = tk.Frame(budget_window)
    scrollable_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    canvas = tk.Canvas(scrollable_frame)
    scrollbar = tk.Scrollbar(scrollable_frame, orient="vertical", command=canvas.yview)
    scrollable_content = tk.Frame(canvas)

    # Configure the canvas and scrollbar
    scrollable_content.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    canvas.create_window((0, 0), window=scrollable_content, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Enable mouse wheel scrolling
    def on_mouse_wheel(event):
        canvas.yview_scroll(-1 * (event.delta // 120), "units")

    canvas.bind_all("<MouseWheel>", on_mouse_wheel)  # For Windows
    canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))  # For Linux
    canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))  # For Linux

    # Create input fields for each category except "Total"
    input_fields = {}  # Dictionary to store input fields for each category
    try:
        for _, row in df.iterrows():
            category = row["Category"]

            # Skip the "Total" category
            if category.lower() == "total":
                continue

            # Create a label for the category
            category_label = tk.Label(scrollable_content, text=f"{category}:", font=("Arial", 12))
            category_label.pack(anchor="w", padx=5, pady=2)

            # Create an input field for the category
            category_entry = tk.Entry(scrollable_content, width=20, font=("Arial", 12))
            category_entry.pack(anchor="w", padx=5, pady=2)

            # Store the input field in the dictionary
            input_fields[category] = category_entry

    except Exception as e:
        messagebox.showerror("Error", f"Failed to create input fields: {e}")

    # Add a button to save the budget inputs
    def save_budget():
        budget_data = {}
        for category, entry in input_fields.items():
            value = entry.get()
            budget_data[category] = value

        # Display the saved budget data (you can save it to a file or process it further)
        messagebox.showinfo("Budget Saved", f"Budget data saved:\n{budget_data}")

    save_button = tk.Button(budget_window, text="Save Budget", command=save_budget, width=15, height=2)
    save_button.pack(pady=10)

# Create the main application window
root = tk.Tk()
root.title("Finance Master GUI")
root.geometry("500x830")  # Set the window size

# Add a label
label = tk.Label(root, text="Welcome to Finance Master\n-Alexander Wiese-", font=("Calibri", 24))
label.pack(pady=15)

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
button_frame.grid(row=1, column=3, rowspan=3, padx=10, pady=5, sticky="n")

# Add buttons to the button frame
upload_button = tk.Button(button_frame, text="Upload Files", command=upload_files, width=15, height=2)
upload_button.pack(pady=10)

refresh_button = tk.Button(button_frame, text="Refresh", command=update_file_lists, width=15, height=2)
refresh_button.pack(pady=10)

run_button = tk.Button(button_frame, text="Run program", command=run_program, width=15, height=2)
run_button.pack(pady=10)

# Add the Budget Creator button
budget_button = tk.Button(button_frame, text="Budget Creator", command=open_budget_creator, width=15, height=2, state=tk.DISABLED)
budget_button.pack(pady=10)

# Add a progress bar
progress_bar = ttk.Progressbar(root, mode="indeterminate", length=400)
progress_bar.pack(pady=20)

# Initialize file lists
update_file_lists()

# Run the application
root.mainloop()