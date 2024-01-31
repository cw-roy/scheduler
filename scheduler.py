import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# Load employee data from Excel
employee_data = pd.read_excel("team_list.xlsx")

# Set the seed for reproducibility
np.random.seed(42)

def generate_weighted_schedule(employee_data):
    # Weighting algorithm: Assign duties to employees who haven't had duty in the last 4 weeks
    last_assigned = datetime.now() - timedelta(weeks=4)
    available_employees = employee_data[employee_data["Availability"] == "yes"]
    available_employees = available_employees[available_employees["LastAssignment"] <= last_assigned]

    if len(available_employees) == 0:
        # If all employees have been assigned in the last 4 weeks, reset the LastAssignment column
        employee_data["LastAssignment"] = datetime.now()

    # Assign duties to employees with less recent duty
    weights = (last_assigned - available_employees["LastAssignment"]).dt.days
    selected_employee = np.random.choice(available_employees.index, p=weights / weights.sum())

    # Update LastAssignment for the selected employee
    employee_data.at[selected_employee, "LastAssignment"] = datetime.now()

    return selected_employee

def generate_duty_schedule(start_date, num_weeks):
    duty_schedule = []

    for week in range(num_weeks):
        # Get the start and end date for the current week
        week_start = start_date + timedelta(weeks=week)
        week_end = week_start + timedelta(days=4)  # Assuming weekdays only

        # Generate duty pairs for the week
        employee1 = generate_weighted_schedule(employee_data)
        employee2 = generate_weighted_schedule(employee_data)

        # Append the duty schedule for the week
        duty_schedule.append({
            "Week": week + 1,
            "Start Date": week_start.strftime("%Y-%m-%d"),
            "End Date": week_end.strftime("%Y-%m-%d"),
            "Employee 1": employee_data.at[employee1, "Name"],
            "Email 1": employee_data.at[employee1, "Email"],
            "Employee 2": employee_data.at[employee2, "Name"],
            "Email 2": employee_data.at[employee2, "Email"],
        })

    return pd.DataFrame(duty_schedule)

# Set up initial LastAssignment column with current date
employee_data["LastAssignment"] = datetime.now()

# Set the start date and number of weeks for the duty schedule
start_date = datetime.now()
num_weeks = 52

# Generate duty schedule
schedule_df = generate_duty_schedule(start_date, num_weeks)

# Save the duty schedule to an Excel file
schedule_df.to_excel("on_point_schedule.xlsx", index=False)

print("Duty schedule generated and saved to 'on_point_schedule.xlsx'")
