import pandas as pd
from datetime import datetime
import os
import tkinter as tk
from tkinter import filedialog

def main():
    # Set up a root window but keep it hidden
    root = tk.Tk()
    root.withdraw()

    # Prompt the user to select the input file
    input_file_path = filedialog.askopenfilename(title="Select the Geezeo Monthly Engagement Report NO CSV", filetypes=[("Excel files", "*.xlsx")])
    if not input_file_path:
        print("No input file selected. Exiting.")
        return

    # Prompt the user to select the output directory
    output_path = filedialog.askdirectory(title="Select where the xlsx and csv purge files will go")
    if not output_path:
        print("No output directory selected. Exiting.")
        return
    # Ensure the path ends with a slash
    output_path += '/' if not output_path.endswith('/') else ''

    # Load the Excel document
    df = pd.read_excel(input_file_path, engine='openpyxl')

    # Apply the formula to determine if a user is active
    df['Status'] = df.apply(lambda row: "Yes" if (row['Last Login Date'] >= pd.Timestamp(datetime.now()).to_period('M').to_timestamp() - pd.DateOffset(months=3)) or (row['Budgets'] >= 1) or (row['Alerts'] >= 1) else "No", axis=1)

    # Filter out the "No" values
    inactive_users = df[df['Status'] == 'No']

    # Define the date format for the output files
    current_time = datetime.now().strftime('%Y%m%d%H')

    # Save to xlsx and CSV with the specified naming format
    inactive_users.to_excel(output_path + f"Purge list processed {current_time}.xlsx", index=False, engine='openpyxl')
    inactive_users.to_csv(output_path + f"Purge list processed {current_time}.csv", index=False)

    print("Files saved successfully!")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"An error occurred: {e}")
