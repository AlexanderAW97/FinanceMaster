import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import subprocess
import webbrowser
import threading
import pandas as pd
from categories import categories  # Import categories from categories.py

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
    budget_window = tk.Toplevel(root)
    budget_window.title("Budget Creator")
    budget_window.geometry("800x800")

    # Add a label
    tk.Label(budget_window, text="Budget Creator", font=("Arial", 18)).pack(pady=10)

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

        # Clear the Treeview before inserting new data (if needed)
        for item in tree.get_children():
            tree.delete(item)

        # Insert data into the Treeview
        for _, row in df.iterrows():
            tree.insert("", tk.END, values=(row["Category"], row["Ut fra konto"], row["Inn på konto"]))

    except Exception as e:
        messagebox.showerror("Error", f"Failed to load budget data: {e}")

    # Add input fields for each category
    scrollable_frame = create_scrollable_frame(budget_window)
    input_fields = create_category_input_fields(scrollable_frame)

    # Add buttons
    button_frame = tk.Frame(budget_window)
    button_frame.pack(pady=10)

    tk.Button(button_frame, text="Save Budget", command=lambda: save_budget(input_fields), width=15, height=2).pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="Add Category", command=open_category_manager, width=15, height=2).pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="Back", command=budget_window.destroy, width=15, height=2).pack(side=tk.LEFT, padx=5)

def create_scrollable_frame(parent):
    """Create a scrollable frame."""
    scrollable_frame = tk.Frame(parent)
    scrollable_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    canvas = tk.Canvas(scrollable_frame)
    scrollbar = tk.Scrollbar(scrollable_frame, orient="vertical", command=canvas.yview)
    content_frame = tk.Frame(canvas)

    content_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    canvas.create_window((0, 0), window=content_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    return content_frame

def create_category_input_fields(parent):
    """Create input fields for each category."""
    input_fields = {}
    for row_index, (category, keywords) in enumerate(categories.items()):
        tk.Label(parent, text=f"{category}:", font=("Arial", 12)).grid(row=row_index, column=0, padx=5, pady=5, sticky="w")
        entry = tk.Entry(parent, width=20, font=("Arial", 12))
        entry.grid(row=row_index, column=1, padx=5, pady=5, sticky="w")
        input_fields[category] = entry
    return input_fields

def save_budget(input_fields):
    """Save the budget data entered by the user."""
    budget_data = {category: entry.get() for category, entry in input_fields.items()}
    messagebox.showinfo("Budget Saved", f"Budget data saved:\n{budget_data}")

def open_category_manager():
    """Open the Category Manager window."""
    category_window = tk.Toplevel(root)
    category_window.title("Category Manager")
    category_window.geometry("800x600")

    # Add a label
    tk.Label(category_window, text="Manage Categories", font=("Arial", 18)).pack(pady=10)

    # Create a scrollable frame for categories
    scrollable_frame = create_scrollable_frame(category_window)
    category_entries = {}

    # Display existing categories
    for row_index, (category, keywords) in enumerate(categories.items()):
        tk.Label(scrollable_frame, text=f"{category}:", font=("Arial", 12)).grid(row=row_index, column=0, padx=5, pady=5, sticky="w")
        text_widget = tk.Text(scrollable_frame, width=50, height=1, font=("Arial", 12), wrap="word")
        text_widget.insert("1.0", ", ".join(keywords))
        text_widget.grid(row=row_index, column=1, padx=5, pady=5, sticky="w")
        category_entries[category] = text_widget

    # Add save button
    tk.Button(category_window, text="Save Changes", command=lambda: save_categories(category_entries), width=15, height=2).pack(pady=10)

def save_categories(category_entries):
    """Save updated categories to categories.py."""
    updated_categories = {category: text_widget.get("1.0", "end-1c").split(",") for category, text_widget in category_entries.items()}
    with open("categories.py", "w", encoding="utf-8") as f:
        f.write("categories = {\n")
        for category, keywords in updated_categories.items():
            f.write(f'    "{category}": {keywords},\n')
        f.write("}\n")
    messagebox.showinfo("Success", "Categories updated successfully!")

def open_uncategorized_manager():
    """Open the Uncategorized Manager window."""
    # Create a new window
    uncategorized_window = tk.Toplevel(root)
    uncategorized_window.title("Uncategorized Manager")
    uncategorized_window.geometry("800x600")

    # Add a label
    label = tk.Label(uncategorized_window, text="Manage Uncategorized Entries", font=("Arial", 18))
    label.pack(pady=10)

    # Create a scrollable frame for the uncategorized entries
    scrollable_frame = tk.Frame(uncategorized_window)
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

    # Load uncategorized entries from combined_output.csv
    combined_output_file = os.path.join(output_folder, "combined_output.csv")
    if not os.path.exists(combined_output_file):
        messagebox.showerror("Error", "combined_output.csv not found. Please run the program first.")
        return

    df = pd.read_csv(combined_output_file)
    uncategorized_df = df[df["Category"] == "Uncategorized"]

    if uncategorized_df.empty:
        messagebox.showinfo("No Uncategorized Entries", "All entries are categorized!")
        uncategorized_window.destroy()
        return

    # Load categories from categories.py
    from categories import categories
    category_names = list(categories.keys())

    # Dictionary to store dropdown selections
    dropdown_selections = {}

    # Display uncategorized entries with dropdowns
    for row_index, (_, row) in enumerate(uncategorized_df.iterrows()):
        # Display the entry description
        description_label = tk.Label(scrollable_content, text=row["Forklaring"], font=("Arial", 12))
        description_label.grid(row=row_index, column=0, padx=5, pady=5, sticky="w")

        # Create a dropdown menu for categories
        selected_category = tk.StringVar(scrollable_content)
        selected_category.set("Select Category")  # Default value
        dropdown = tk.OptionMenu(scrollable_content, selected_category, *category_names)
        dropdown.grid(row=row_index, column=1, padx=5, pady=5, sticky="w")

        # Store the dropdown selection
        dropdown_selections[row_index] = (row, selected_category)

    # Save changes to combined_output.csv
    def save_changes():
        for row_index, (row, selected_category) in dropdown_selections.items():
            new_category = selected_category.get()
            if new_category != "Select Category":
                df.loc[df.index == row.name, "Category"] = new_category

        # Save the updated DataFrame back to combined_output.csv
        df.to_csv(combined_output_file, index=False)
        messagebox.showinfo("Success", "Uncategorized entries updated successfully!")
        uncategorized_window.destroy()

    # Add a save button
    save_button = tk.Button(uncategorized_window, text="Save Changes", command=save_changes, width=15, height=2)
    save_button.pack(pady=10)

# Create the main application window
root = tk.Tk()
root.title("Finance Master GUI")
root.geometry("500x830")  # Set the window size

# Add a label
label = tk.Label(root, text="Welcome to FinanceMaster", font=("Calibri", 24))
sublabel = tk.Label(root, text="Made by Alexander Wiese", font=("Calibri", 14))
label.pack(pady=20)
sublabel.pack(pady=5)

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

# Add the "Add Category" button
add_category_button = tk.Button(button_frame, text="Add Category", command=open_category_manager, width=15, height=2)
add_category_button.pack(pady=10)

# Add a progress bar
progress_bar = ttk.Progressbar(root, mode="indeterminate", length=400)
progress_bar.pack(pady=20)
 
# Initialize file lists
update_file_lists()

# Run the application
root.mainloop()