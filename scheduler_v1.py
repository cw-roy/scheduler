import pandas as pd
from datetime import datetime, timedelta

# Load team data from Excel
team_file_path = "team_list.xlsx"
team_data = pd.read_excel(team_file_path)

# Define the duration of the duty assignment (one week)
assignment_duration = timedelta(days=5)  # Monday to Friday

# Function to generate duty assignments
def generate_assignments(team_data):
    assignments = []

    # Initialize start date
    start_date = datetime(2024, 1, 1).replace(hour=0, minute=0, second=0, microsecond=0)

    # Iterate through pairs of employees
    for i in range(0, len(team_data), 2):
        employee1 = team_data.iloc[i]['Name']
        employee2 = team_data.iloc[i + 1]['Name']

        # Specify start and end dates for the assignment (Monday to Friday)
        end_date = start_date + assignment_duration - timedelta(days=1)

        # Append assignment details to the list
        assignments.append({'Pair': f"{employee1}, {employee2}", 'Start': start_date, 'End': end_date})

        # Update start date for the next pair
        start_date = start_date + timedelta(weeks=1)

    return pd.DataFrame(assignments)

# Save assignments to Excel
assignments_df = generate_assignments(team_data)
output_file_path = "scheduler_v1_assignments.xlsx"

# Convert datetime columns to formatted strings
assignments_df['Start'] = assignments_df['Start'].dt.strftime('%m-%d-%Y')
assignments_df['End'] = assignments_df['End'].dt.strftime('%m-%d-%Y')

# Save to Excel
assignments_df.to_excel(output_file_path, index=False)

print(f"Assignments saved to {output_file_path}")
