import pandas as pd
from datetime import datetime
import os

def main():
    # Define the paths for input and output
    input_path = "C:\\Users\\WhatA\\Documents\\Geezeo Monthly Engagement Report\\"
    output_path = "C:\\Users\\WhatA\\Documents\\Geezeo Monthly Engagement Report\\Output\\"

    # Specify the input file name
    input_file = "Monthly_Engagement_Report_2023-10-02.xlsx"

    # Load the Excel document
    df = pd.read_excel(input_path + input_file, engine='openpyxl')

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
