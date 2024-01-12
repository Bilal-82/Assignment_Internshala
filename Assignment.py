import pandas as pd

def analyze_spreadsheet(file_path):
    # This command will load the spreadsheet into a pandas DataFrame
    df = pd.read_excel(file_path)

    # As the 'Time' and 'Time Out' columns are in datetime format
    df['Time'] = pd.to_datetime(df['Time'])
    df['Time Out'] = pd.to_datetime(df['Time Out'])

    # Sorting the DataFrame by 'Employee Name' and 'Time'
    df.sort_values(['Employee Name', 'Time'], inplace=True)

    # Initialize variables for consecutive days and previous time out
    consecutive_days_count = 0
    prev_time_out = None

    # Initialize lists to store names for each condition
    consecutive_days_names = []
    less_than_10_hours_names = []
    more_than_14_hours_names = []

    # Iterate through the DataFrame to analyze the data
    for _, row in df.iterrows():
        # Solving for each part one by one.
        # a) 7 consecutive days
        if consecutive_days_count == 6:
            consecutive_days_names.append(row['Employee Name'])
            consecutive_days_count = 0
        elif prev_time_out is not None and (row['Time'] - prev_time_out).days <= 1:
            consecutive_days_count += 1
        else:
            consecutive_days_count = 0

        # b) Less than 10 hours between shifts but greater than 1 hour
        if prev_time_out is not None and (row['Time'] - prev_time_out).total_seconds() / 3600 < 10 and (row['Time'] - prev_time_out).total_seconds() / 3600 > 1:
            less_than_10_hours_names.append(row['Employee Name'])

        # c) More than 14 hours in a single shift
        if (row['Time Out'] - row['Time']).total_seconds() / 3600 > 14:
            more_than_14_hours_names.append(row['Employee Name'])

        # Update previous time out
        prev_time_out = row['Time Out']

    # Print or return the collected names with line breaks
    print("Employees with 7 consecutive days:\n", '\n'.join(consecutive_days_names))
    print("\nEmployees with less than 10 hours between shifts but greater than 1 hour:\n", '\n'.join(less_than_10_hours_names))
    print("\nEmployees with more than 14 hours in a single shift:\n", '\n'.join(more_than_14_hours_names))

if __name__ == "__main__":
    file_path = 'Assignment_Timecard.xlsx'  # Replace with the actual file path
    analyze_spreadsheet(file_path)
