import os
import pandas as pd
from difflib import SequenceMatcher
import openpyxl
from openpyxl.styles import Font
from categories import categories  # Import categories from the separate file

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
    category_column = "Category"
    df[category_column] = "Uncategorized"

    for category, keywords in categories.items():
        mask = (df[category_column] == "Uncategorized") & df[column_name].str.contains('|'.join(keywords), case=False, na=False)
        df.loc[mask, category_column] = category

    return df

# Function to create budget Excel file
def create_budget_excel(output_file, df):
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

    # Generate budget Excel file
    budget_output_file = os.path.join(output_folder, "Totals.xlsx")
    create_budget_excel(budget_output_file, combined_df)