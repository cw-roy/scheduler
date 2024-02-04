import os
import pandas as pd
import random
from datetime import datetime, timedelta
import logging
from collections import deque, defaultdict

# Configure logging
log_directory = os.path.join(os.path.dirname(__file__), "logging")
os.makedirs(log_directory, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        # logging.StreamHandler(),  # Comment out to stop logging to console
        logging.FileHandler(
            os.path.join(log_directory, "scheduler_events.log")
        )  # Log to file
    ],
)


# Log the start of the script execution
logging.info("Script execution started.")

# Define the assignment data log file
assignment_data_log_path = os.path.join(log_directory, "assignment_data_log.csv")


def initialize_assignment_data_log():
    # Initialize the assignment data log file
    with open(assignment_data_log_path, "w") as log_file:
        log_file.write(
            "Tech,Number of Assignments,Last Assignment Date,Workload History\n"
        )


def log_assignment_data(assignment_data):
    # Log assignment data to the assignment_data_log file
    with open(assignment_data_log_path, "a") as log_file:
        for tech, data in assignment_data.items():
            log_file.write(
                f"{tech},{data['num_assignments']},{data['last_assignment_date']},{','.join(data['workload_history'])}\n"
            )


def backup_existing_assignments():
    file_name = "assignments.xlsx"
    history_directory = "history"

    working_directory = "working"
    file_path = os.path.join(working_directory, file_name)

    if os.path.exists(file_path):
        timestamp_str = datetime.now().strftime("%m-%d-%Y_%H-%M-%S")
        new_file_name = os.path.join(
            history_directory, f"assignments_as_of_{timestamp_str}.xlsx"
        )

        os.rename(file_path, new_file_name)
        logging.info("Assignment file backed up to {}".format(new_file_name))
        print(f"Existing {file_name} backed up to {new_file_name}")
    else:
        logging.info("No existing assignments to back up.")


def read_employee_data(file_path):
    """Read employee data from the Excel spreadsheet."""
    try:
        # df = pd.read_excel(file_path)
        df = pd.read_excel(os.path.join("working", file_path))

        # Clean column names by removing leading and trailing spaces
        df.columns = df.columns.str.strip()

        # Validate columns
        expected_columns = ["Name", "Email", "Available"]
        actual_columns_stripped = [col.strip() for col in df.columns]

        if set(actual_columns_stripped) != set(expected_columns):
            raise ValueError(
                f"Expected Columns: {expected_columns}\nActual Columns: {actual_columns_stripped}"
            )

        return df

    except Exception as e:
        logging.error(f"Error reading employee data: {e}")
        return None


def get_email_addresses(employee_data, pair):
    """Get email addresses for the given pair of employees."""
    emails = []
    for name in pair:
        email = employee_data.loc[employee_data["Name"] == name, "Email"].values[0]
        emails.append(email)
    return emails


def normalize_weights(weights, max_assignments):
    """Normalize weights to ensure they add up to max_assignments."""
    total_weight = sum(weights)
    return [w / total_weight * max_assignments for w in weights]


def detect_changes(old_data, new_data):
    """Detect changes between two versions of employee data."""
    changes = []

    for index, row in old_data.iterrows():
        old_row = row.to_dict()
        new_row = new_data[new_data["Name"] == old_row["Name"]].to_dict("records")[0]

        for key, old_value in old_row.items():
            new_value = new_row[key]
            if old_value != new_value:
                changes.append(
                    f"Change in {key} for employee {old_row['Name']}: {old_value} -> {new_value}"
                )

    return changes


def calculate_max_assignments(weeks_in_year, available_tech_count):
    return weeks_in_year / available_tech_count

def initialize_assigned_pairs_queue(size):
    """Initialize the assigned pairs queue with the specified size."""
    return deque(maxlen=size)

def initialize_weights(employee_data):
    """Initialize weights based on the number of available employees."""
    return [1.0] * len(employee_data[employee_data["Available"] == "yes"])

def generate_rotation_schedule(employee_data, weeks_in_year):
    schedule = {}

    current_date = datetime.now()
    assigned_pairs_queue = initialize_assigned_pairs_queue(4)
    max_assignments = calculate_max_assignments(
        weeks_in_year, len(employee_data[employee_data["Available"] == "yes"])
    )

    # Load the previous employee data
    previous_employee_data = read_employee_data("team_list.xlsx")

    if previous_employee_data is not None:
        # Detect and log changes in employee data
        changes = detect_changes(previous_employee_data, employee_data)
        if changes:
            logging.info("Changes detected in team_list.xlsx:")
            for change in changes:
                logging.info(change)

    # Initialize assignment data dictionary
    assignment_data = defaultdict(
        lambda: {
            "num_assignments": 0,
            "last_assignment_date": "",
            "workload_history": [],
        }
    )

    weights = initialize_weights(employee_data)

    for week in range(1, 53):
        start_date, end_date = calculate_week_dates(current_date, week)

        paired_employees = generate_paired_employees(
            employee_data, weights, assigned_pairs_queue, max_assignments
        )

        weights = update_weights(weights)

        logging.debug("Weights after Week {} Assignment: {}".format(week, weights))

        # Update assignment data
        update_assignment_data(assignment_data, paired_employees, end_date, start_date)

        schedule[week] = {
            "start_date": start_date.strftime("%m-%d-%Y"),
            "end_date": end_date.strftime("%m-%d-%Y"),
            "pair": paired_employees,
            "email_addresses": get_email_addresses(employee_data, paired_employees),
        }

    # Log assignment data at the end of script execution
    log_assignment_data(assignment_data)

    return schedule

def calculate_week_dates(current_date, week):
    """Calculate the start and end dates of a week."""
    start_date = current_date + timedelta(
        days=((week - 1) * 7) + (0 - current_date.weekday()) % 7
    )
    end_date = start_date + timedelta(days=4)
    return start_date, end_date

def generate_paired_employees(employee_data, weights, assigned_pairs_queue, max_assignments):
    """Generate a pair of employees for assignment."""
    paired_employees = None
    while True:
        normalized_weights = normalize_weights(weights, max_assignments)
        paired_employees = random.choices(
            employee_data[employee_data["Available"] == "yes"]["Name"].tolist(),
            weights=normalized_weights,
            k=2,
        )
        if all(pair not in assigned_pairs_queue for pair in paired_employees):
            break
    assigned_pairs_queue.extend(paired_employees)
    return paired_employees

def update_weights(weights):
    """Update weights based on a random factor."""
    return [w * random.uniform(0.8, 1.2) for w in weights]

def update_assignment_data(assignment_data, paired_employees, end_date, start_date):
    """Update assignment data for each employee."""
    for tech in paired_employees:
        assignment_data[tech]["num_assignments"] += 1
        assignment_data[tech]["last_assignment_date"] = end_date.strftime(
            "%m-%d-%Y"
        )
        assignment_data[tech]["workload_history"].append(
            start_date.strftime("%m-%d-%Y")
        )



# def generate_rotation_schedule(employee_data, weeks_in_year):
#     schedule = {}
#     weights = [1.0] * len(employee_data[employee_data["Available"] == "yes"])

#     current_date = datetime.now()
#     assigned_pairs_queue = deque(maxlen=4)
#     max_assignments = calculate_max_assignments(
#         weeks_in_year, len(employee_data[employee_data["Available"] == "yes"])
#     )

#     # Load the previous employee data
#     previous_employee_data = read_employee_data("team_list.xlsx")

#     if previous_employee_data is not None:
#         # Detect and log changes in employee data
#         changes = detect_changes(previous_employee_data, employee_data)
#         if changes:
#             logging.info("Changes detected in team_list.xlsx:")
#             for change in changes:
#                 logging.info(change)

#     # Initialize assignment data dictionary
#     assignment_data = defaultdict(
#         lambda: {
#             "num_assignments": 0,
#             "average_frequency": 0,
#             "last_assignment_date": "",
#             "workload_history": [],
#         }
#     )

#     for week in range(1, 53):
#         start_date = current_date + timedelta(
#             days=((week - 1) * 7) + (0 - current_date.weekday()) % 7
#         )
#         end_date = start_date + timedelta(days=4)

#         paired_employees = None
#         while True:
#             normalized_weights = normalize_weights(weights, max_assignments)
#             paired_employees = random.choices(
#                 employee_data[employee_data["Available"] == "yes"]["Name"].tolist(),
#                 weights=normalized_weights,
#                 k=2,
#             )
#             if all(pair not in assigned_pairs_queue for pair in paired_employees):
#                 break

#         assigned_pairs_queue.extend(paired_employees)

#         weights = [w * random.uniform(0.8, 1.2) for w in weights]

#         logging.debug("Weights after Week {} Assignment: {}".format(week, weights))

#         # Update assignment data
#         for tech in paired_employees:
#             assignment_data[tech]["num_assignments"] += 1
#             assignment_data[tech]["last_assignment_date"] = end_date.strftime(
#                 "%m-%d-%Y"
#             )
#             assignment_data[tech]["workload_history"].append(
#                 start_date.strftime("%m-%d-%Y")
#             )

#         schedule[week] = {
#             "start_date": start_date.strftime("%m-%d-%Y"),
#             "end_date": end_date.strftime("%m-%d-%Y"),
#             "pair": paired_employees,
#             "email_addresses": get_email_addresses(employee_data, paired_employees),
#         }

#     # Log assignment data at the end of script execution
#     log_assignment_data(assignment_data)

#     return schedule


def write_to_excel(schedule):
    """Write the schedule to an Excel file with the specified format."""
    flat_schedule = []

    for week, data in schedule.items():
        start_date = datetime.strptime(data["start_date"], "%m-%d-%Y")
        end_date = datetime.strptime(data["end_date"], "%m-%d-%Y")

        flat_schedule.append(
            {
                "start_date": start_date.strftime("%m-%d-%Y"),
                "end_date": end_date.strftime("%m-%d-%Y"),
                "Agent 1": data["pair"][0],
                "Email1": data["email_addresses"][0],
                "Agent 2": data["pair"][1],
                "Email2": data["email_addresses"][1],
            }
        )

    df = pd.DataFrame(flat_schedule)

    # Specify the path within the "working" directory
    working_directory = "working"
    output_file_path = os.path.join(working_directory, "assignments.xlsx")

    # Write to Excel file with the specified columns
    df.to_excel(output_file_path, index=False)


def main():
    employee_data = read_employee_data("team_list.xlsx")

    if employee_data is not None:
        backup_existing_assignments()

        # Initialize the assignment data log file
        initialize_assignment_data_log()

        rotation_schedule = generate_rotation_schedule(employee_data, weeks_in_year=52)

        # For production use, write the schedule to an Excel file
        write_to_excel(rotation_schedule)

    else:
        logging.error("Exiting program due to errors.")

    # Log the end of the script execution
    logging.info("Script execution completed.\n")


if __name__ == "__main__":
    main()
