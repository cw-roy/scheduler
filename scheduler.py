import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# Function to load employee data from an Excel file
def load_employee_data(file_path):
    return pd.read_excel(file_path)

# Function to save the duty schedule to an Excel file
def save_duty_schedule(schedule_df, file_path):
    schedule_df.to_excel(file_path, index=False)

# Function to generate a weighted schedule
def generate_weighted_schedule(employee_data, last_assigned):
    available_employees = employee_data[employee_data["Availability"] == "yes"]

    # Identify employees who haven't had duty in the last four weeks
    not_assigned_last_4_weeks = available_employees["LastAssignment"] <= last_assigned - timedelta(weeks=4)
    available_employees = available_employees[not_assigned_last_4_weeks]

    # If there are no such employees, consider all available employees
    if len(available_employees) == 0:
        available_employees = employee_data[employee_data["Availability"] == "yes"]

    weights = (last_assigned - available_employees["LastAssignment"]).dt.days
    weights = np.maximum(weights, 0)  # Ensure probabilities are non-negative

    # Handle the case where all weights are zero
    if weights.sum() == 0:
        selected_employee = np.random.choice(available_employees.index)
    else:
        selected_employee = np.random.choice(available_employees.index, p=weights / weights.sum())

    # Update LastAssignment for the selected employee
    employee_data.at[selected_employee, "LastAssignment"] = last_assigned

    return selected_employee

# Function to get a unique pair of employees
def get_unique_pair(employee_data, last_assigned, selected_indices):
    # Shuffle the available employees
    available_employees = employee_data[employee_data["Availability"] == "yes"].sample(frac=1)

    # Filter out employees with recent assignments
    available_employees = available_employees[~available_employees.index.isin(selected_indices)]

    if len(available_employees) >= 2:
        # Take the first two employees for duty
        employee1 = available_employees.iloc[0]
        employee2 = available_employees.iloc[1]

        # Update LastAssignment for the selected employees
        employee_data.at[employee1.name, "LastAssignment"] = last_assigned
        employee_data.at[employee2.name, "LastAssignment"] = last_assigned

        return employee1, employee2
    else:
        return None, None

# Function to generate the duty schedule
def generate_duty_schedule(employee_data, start_date, num_weeks):
    duty_schedule = []
    available_employees = employee_data[employee_data["Availability"] == "yes"]
    selected_indices = set()

    for week in range(num_weeks):
        # Get the start and end date for the current week
        week_start = start_date + timedelta(weeks=week)
        week_end = week_start + timedelta(days=4)  # Assuming weekdays only

        # Check if we have enough available employees for the entire week
        if len(available_employees) >= 2:
            # Sample two unique employees from the available pool
            selected_employees = available_employees.sample(n=2)

            # Update LastAssignment for the selected employees
            for _, employee in selected_employees.iterrows():
                selected_indices.add(employee.name)
                employee_data.at[employee.name, "LastAssignment"] = week_end

            # Append the duty schedule for the week
            duty_schedule.append({
                "Week": week + 1,
                "Start Date": week_start.strftime("%m-%d-%Y"),
                "End Date": week_end.strftime("%m-%d-%Y"),
                "Employee 1": selected_employees.iloc[0]["Name"],
                # "Email 1": selected_employees.iloc[0]["Email"],
                "Employee 2": selected_employees.iloc[1]["Name"],
                # "Email 2": selected_employees.iloc[1]["Email"],
            })
        else:
            # Handle the case where there are not enough available employees
            print(f"Not enough available employees for duty in week {week + 1}")
            break

    return pd.DataFrame(duty_schedule)

# Main script execution
if __name__ == "__main__":
    # Load employee data from Excel
    employee_data = load_employee_data("team_list.xlsx")

    # Set up initial LastAssignment column with the current date
    employee_data["LastAssignment"] = datetime.now()

    # Set the start date and number of weeks for the duty schedule
    start_date = datetime.now()
    num_weeks = 52

    # Generate duty schedule
    schedule_df = generate_duty_schedule(employee_data, start_date, num_weeks)

    # Save the duty schedule to an Excel file
    save_duty_schedule(schedule_df, "on_point_schedule.xlsx")

    print("Duty schedule generated and saved to 'on_point_schedule.xlsx'")
