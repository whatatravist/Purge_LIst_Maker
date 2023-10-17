import pandas as pd
from datetime import datetime
import os

def main():
    # File paths
    input_path = "C:\\Users\\ttedford\\Documents\\Geezeo Monthly Engagement Report\\"
    output_path = "C:\\Users\\ttedford\\Documents\\"

    # Identify the file with date format YYYY-MM-DD in its name
    for file in os.listdir(input_path):
        if "Monthly_Engagement_Report" in file and "-" in file:
            input_file = file

    # Load the Excel document
    df = pd.read_excel(input_path + input_file, engine='openpyxl')

    # Apply the formula to determine if a user is active
    df['Status'] = df.apply(lambda row: "Yes" if (row['Last Login Date'] >= pd.Timestamp(datetime.now()).to_period('M').to_timestamp() - pd.DateOffset(months=3)) or (row['Budgets'] >= 1) else "No", axis=1)

    # Filter out the "No" values
    inactive_users = df[df['Status'] == 'No']

    # Extract date from input filename
    date_str = input_file.split(' ')[3]  # assuming format like "prefix YYYY-MM-DD suffix.xlsx"

    # Save to xlsx and CSV
    inactive_users.to_excel(output_path + f"Purge List {date_str} XLSX.xlsx", index=False, engine='openpyxl')
    inactive_users.to_csv(output_path + f"Purge List {date_str} CSV.csv", index=False)

    print("Files saved successfully!")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"An error occurred: {e}")

