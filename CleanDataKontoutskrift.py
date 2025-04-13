import os
import pandas as pd
from difflib import SequenceMatcher
import openpyxl
from openpyxl.styles import Font
from categories import categories  # Import categories from the separate file
import tkinter as tk
from tkinter import ttk, messagebox
import re

# Define paths
base_folder = os.path.dirname(os.path.abspath(__file__))
input_folder = os.path.join(base_folder, "InputFolder")
output_folder = os.path.join(base_folder, "OutputFolder")

# Ensure folders exist
os.makedirs(output_folder, exist_ok=True)

# Column names
column_dato = "Dato"
column_forklaring = "Forklaring"
column_rentedato = "Rentedato"
column_ut_fra_konto = "Ut fra konto"
column_inn_pa_konto = "Inn på konto"

# Combine output flag
combine_output = 1  # Set to 0 for separate files per input

# Function to find similar names
def find_similar_names(df, column_name, similarity_threshold=0.8):
    unique_names = df[column_name].unique()
    similar_groups = {}

    for i, name1 in enumerate(unique_names):
        for name2 in unique_names[i + 1:]:
            similarity = SequenceMatcher(None, name1, name2).ratio()
            if similarity >= similarity_threshold:
                if name1 not in similar_groups:
                    similar_groups[name1] = set()
                similar_groups[name1].add(name2)

    for base_name, similar_names in similar_groups.items():
        for similar_name in similar_names:
            df[column_name] = df[column_name].replace(similar_name, base_name)

# Function to categorize entries
def categorize_entries(df, column_name, categories):
    """
    Categorize entries in the DataFrame based on keywords in the specified column.
    Appends an underscore and iteration number to each category entry.
    """
    category_column = "Category"
    df[category_column] = "Uncategorized"

    # Create a dictionary to track counts for each category
    category_counts = {category: 0 for category in categories.keys()}

    for category, keywords in categories.items():
        # Create a regex pattern to match whole words for each keyword
        pattern = r'\b(?:' + '|'.join(re.escape(keyword) for keyword in keywords) + r')\b'

        # Create a mask for rows that match the category
        mask = df[column_name].str.contains(pattern, case=False, na=False)

        # Update the category column with unique category names
        for index in df[mask].index:
            # Increment the count for the category and assign a unique name
            category_counts[category] += 1
            df.at[index, category_column] = f"{category}_{category_counts[category]}"

    return df

def clean_category_names(df, category_column="Category"):
    """
    Remove the underscore and iteration number from the category names in the specified column.
    """
    df[category_column] = df[category_column].str.replace(r'_\d+$', '', regex=True)
    return df

# Function to create budget Excel file
def create_budget_excel(output_file, df):
    """
    Create an Excel file summarizing the totals for each category.
    """
    # Clean category names
    df = clean_category_names(df)

    # Aggregate totals
    if "Category" in df.columns and "Ut fra konto" in df.columns:
        ut_fra_konto_totals = df.groupby("Category")["Ut fra konto"].sum().to_dict()
    else:
        raise ValueError("The required columns ('Category' and 'Ut fra konto') are missing.")

    if "Category" in df.columns and "Inn på konto" in df.columns:
        inn_pa_konto_totals = df.groupby("Category")["Inn på konto"].sum().to_dict()
    else:
        raise ValueError("The required columns ('Category' and 'Inn på konto') are missing.")

    # Create a new workbook
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Totals"  # Change the sheet title to "Totals"

    # Add headers
    sheet["A1"] = "Category"
    sheet["B1"] = "Ut fra konto"
    sheet["C1"] = "Inn på konto"
    sheet["A1"].font = Font(bold=True)
    sheet["B1"].font = Font(bold=True)
    sheet["C1"].font = Font(bold=True)

    # Populate the spreadsheet
    all_categories = set(ut_fra_konto_totals.keys()).union(inn_pa_konto_totals.keys())
    total_ut_fra_konto = 0
    total_inn_pa_konto = 0

    for row, category in enumerate(sorted(all_categories), start=2):
        sheet[f"A{row}"] = category
        sheet[f"B{row}"] = ut_fra_konto_totals.get(category, 0)
        sheet[f"C{row}"] = inn_pa_konto_totals.get(category, 0)
        total_ut_fra_konto += ut_fra_konto_totals.get(category, 0)
        total_inn_pa_konto += inn_pa_konto_totals.get(category, 0)

    # Add total row
    total_row = len(all_categories) + 2
    sheet[f"A{total_row}"] = "Total"
    sheet[f"B{total_row}"] = total_ut_fra_konto
    sheet[f"C{total_row}"] = total_inn_pa_konto
    sheet[f"A{total_row}"].font = Font(bold=True)
    sheet[f"B{total_row}"].font = Font(bold=True)
    sheet[f"C{total_row}"].font = Font(bold=True)

    # Save the workbook
    workbook.save(output_file)
    print(f"Totals spreadsheet saved to {output_file}")

# Function to open the Budget Creator window
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

    # Load categories from categories.py
    from categories import categories

    # Create input fields for each category
    input_fields = {}  # Dictionary to store input fields for each category
    for row_index, (category, keywords) in enumerate(categories.items()):
        # Create a label for the category
        category_label = tk.Label(scrollable_content, text=f"{category}:", font=("Arial", 12))
        category_label.grid(row=row_index, column=0, padx=5, pady=5, sticky="w")

        # Create an input field for the category
        category_entry = tk.Entry(scrollable_content, width=20, font=("Arial", 12))
        category_entry.grid(row=row_index, column=1, padx=5, pady=5, sticky="w")

        # Store the input field in the dictionary
        input_fields[category] = category_entry

    # Add a button to save the budget inputs
    def save_budget():
        budget_data = {}
        for category, entry in input_fields.items():
            value = entry.get()
            budget_data[category] = value

        # Display the saved budget data (you can save it to a file or process it further)
        messagebox.showinfo("Budget Saved", f"Budget data saved:\n{budget_data}")

    # Add a frame for the buttons
    button_frame = tk.Frame(budget_window)
    button_frame.pack(pady=10)

    # Add the "Save Budget" button
    save_button = tk.Button(button_frame, text="Save Budget", command=save_budget, width=15, height=2)
    save_button.pack(side=tk.LEFT, padx=5)

    # Add the "Back" button
    back_button = tk.Button(button_frame, text="Back", command=budget_window.destroy, width=15, height=2)
    back_button.pack(side=tk.LEFT, padx=5)

# Main processing
combined_df = pd.DataFrame()

for filename in os.listdir(input_folder):
    if filename.endswith(".xlsx"):
        input_file_path = os.path.join(input_folder, filename)
        output_file = os.path.join(output_folder, f"cleaned_{os.path.splitext(filename)[0]}.csv")

        df = pd.read_excel(input_file_path, engine="openpyxl")
        df.columns = df.columns.str.strip()
        df = df.iloc[:, 1:]
        for col in df.select_dtypes(include=["object"]).columns:
            df[col] = df[col].str.strip()
        df[column_ut_fra_konto] = pd.to_numeric(df[column_ut_fra_konto], errors="coerce").fillna(0)
        df[column_inn_pa_konto] = pd.to_numeric(df[column_inn_pa_konto], errors="coerce").fillna(0)

        df = categorize_entries(df, column_forklaring, categories)
        find_similar_names(df, column_forklaring)

        # Filter out "Kontooverføringer" category
        df = df[df["Category"] != "Kontooverføringer"]

        if combine_output:
            combined_df = pd.concat([combined_df, df], ignore_index=True)
        else:
            df.to_csv(output_file, index=False, sep=",", encoding="utf-8")
            print(f"Processed and saved: {output_file}")

if combine_output:
    combined_output_file = os.path.join(output_folder, "combined_output.csv")
    combined_df.to_csv(combined_output_file, index=False, sep=",", encoding="utf-8")
    print(f"Processed and saved combined output: {combined_output_file}")

    # Clean category names before generating the budget Excel file
    budget_output_file = os.path.join(output_folder, "Totals.xlsx")
    combined_df = clean_category_names(combined_df)
    create_budget_excel(budget_output_file, combined_df)

# Create the main window
root = tk.Tk()
root.title("Main Window")
root.geometry("400x400")

# Add a button to open the Budget Creator
open_budget_button = tk.Button(root, text="Open Budget Creator", command=open_budget_creator, width=20, height=2)
open_budget_button.pack(pady=20)

# Run the main loop
root.mainloop()