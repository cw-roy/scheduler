import os
import pandas as pd
import random
from datetime import datetime, timedelta
import logging
from collections import deque, defaultdict
import sys

# Configure logging
log_directory = os.path.join(os.path.dirname(__file__), "logging")
os.makedirs(log_directory, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%m-%d-%Y %H:%M:%s",
    handlers=[
        # logging.StreamHandler(),  # Comment out to stop logging to console
        logging.FileHandler(
            os.path.join(log_directory, "scheduler_events.log")
        )  # Log to file
    ],
)

# Set the paths to the input and output files
team_list_path = os.path.join("working", "team_list.xlsx")
assignments_path = os.path.join("working", "assignments.xlsx")


def read_employee_data(file_path):
    """Read employee data from the Excel spreadsheet."""
    try:
        # df = pd.read_excel(file_path)
        df = pd.read_excel(file_path)

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


employee_data = read_employee_data(team_list_path)

# Log the start of the script execution
logging.info("Script execution started.")

# Define the assignment data log file
assignment_data_log_filename = "assignment_data_log.csv"
assignment_data_log_path = os.path.join(
    os.path.dirname(__file__), "logging", assignment_data_log_filename
)

# Check if the assignment_data_log.csv file exists, create it if not
if not os.path.exists(assignment_data_log_path):
    with open(assignment_data_log_path, "w") as log_file:
        log_file.write(
            "Run Start, Tech, Number of Assignments, Last Assignment Date, Workload History\n"
        )


# Load previous employee data and detect changes
previous_employee_data = read_employee_data(team_list_path)


def check_and_exit_if_no_changes(current_employee_data, previous_employee_data):
    """Check for changes in employee availability and exit if no changes."""
    if current_employee_data is None or previous_employee_data is None:
        log_activity("Error reading employee data. Exiting script.")
        logging.info("Script execution completed with errors.\n")
        sys.exit(1)

    if current_employee_data.equals(previous_employee_data):
        log_activity("No changes in availability. Exiting script.")
        logging.info("Script execution completed.\n")
        sys.exit(0)


def initialize_assignment_data_log():
    # Check if the assignment_data_log.csv file exists
    if not os.path.exists(assignment_data_log_path):
        with open(assignment_data_log_path, "w") as log_file:
            log_file.write(
                "Run Start, Tech, Number of Assignments, Last Assignment Date, Workload History\n"
            )


def load_assignment_history():
    """Load historical assignment data from assignment_data_log.csv."""
    history = defaultdict(
        lambda: {
            "num_assignments": 0,
            "last_assignment_date": "",
            "workload_history": [],
        }
    )

    try:
        with open(assignment_data_log_path, "r") as log_file:
            next(log_file)  # Skip header
            for line in log_file:
                parts = line.strip().split(",")
                tech = parts[0]
                num_assignments = int(parts[1])
                last_assignment_date = parts[2]
                workload_history = parts[3:]
                history[tech] = {
                    "num_assignments": num_assignments,
                    "last_assignment_date": last_assignment_date,
                    "workload_history": workload_history,
                }
    except FileNotFoundError:
        logging.warning("No assignment history found. Continuing with empty history.")

    return history


def log_assignment_data(assignment_data):
    # Log assignment data to the assignment_data_log file
    with open(assignment_data_log_path, "a") as log_file:
        # Write assignment data for each tech
        for tech, data in assignment_data.items():
            # Check if 'num_assignments' key exists in the data dictionary
            if "num_assignments" in data:
                log_file.write(
                    f"{tech},{data['num_assignments']},{data['last_assignment_date']},{','.join(data['workload_history'])}\n"
                )


def log_activity(activity_description):
    """Log activities to scheduler_events.log."""
    formatted_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    logging.info(f"{formatted_timestamp} - {activity_description}")


def backup_existing_assignments():
    file_name = "assignment_data_log.csv"
    history_directory = "history"

    file_path = os.path.join(log_directory, file_name)

    if os.path.exists(file_path):
        timestamp_str = datetime.now().strftime("%m-%d-%Y_%H-%M-%S")
        new_file_name = os.path.join(
            history_directory, f"{file_name}_as_of_{timestamp_str}.csv"
        )

        os.rename(file_path, new_file_name)
        logging.info("Assignment data log backed up to {}".format(new_file_name))

        # Keep only the last 3 backups, delete older ones
        delete_old_backups(history_directory, file_name, keep_latest=3)
    else:
        logging.info("No existing assignment data log to back up.")


def delete_old_backups(directory, base_name, keep_latest):
    files = [f for f in os.listdir(directory) if f.startswith(f"{base_name}_as_of_")]
    files.sort(key=lambda x: os.path.getctime(os.path.join(directory, x)), reverse=True)

    # Delete older backups, keep only the last 'keep_latest'
    for old_file in files[keep_latest:]:
        os.remove(os.path.join(directory, old_file))
        logging.info(
            f"Retain maximum of 3 saved logs. Old assignment data log backup deleted: {old_file}"
        )


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


def load_and_detect_changes(previous_data_path, current_data, log_changes=True):
    """Load previous employee data, compare with current data, and log changes."""
    previous_employee_data = read_employee_data(previous_data_path)

    if previous_employee_data is not None:
        # Detect and log changes in employee data
        changes = detect_changes(previous_employee_data, current_data)
        if changes and log_changes:
            logging.info("Changes detected in team_list.xlsx:")
            for change in changes:
                logging.info(change)

    return previous_employee_data


def calculate_week_dates(current_date, week):
    """Calculate the start and end dates of a week."""
    start_date = current_date + timedelta(
        days=((week - 1) * 7) + (0 - current_date.weekday()) % 7
    )
    end_date = start_date + timedelta(days=4)
    return start_date, end_date


def generate_paired_employees(
    employee_data, weights, assigned_pairs_queue, max_assignments
):
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


def update_weights(weights, history=None):
    """Update weights based on certain criteria."""
    return adjust_weights(weights, history)


def adjust_weights(weights, history):
    """Adjust weights based on certain criteria and historical data."""
    adjusted_weights = []

    for tech, weight in zip(employee_data["Name"], weights):
        # Consider the historical number of assignments for each tech
        num_assignments = history[tech]["num_assignments"]

        # Modify the weight based on historical data (customize this part)
        adjusted_weight = weight * (1.0 + 0.1 * num_assignments)  # Example adjustment

        adjusted_weights.append(adjusted_weight)

    return adjusted_weights


def update_assignment_data(assignment_data, paired_employees, end_date, start_date):
    """Update assignment data for each employee."""
    for tech in paired_employees:
        assignment_data[tech]["num_assignments"] += 1
        assignment_data[tech]["last_assignment_date"] = end_date.strftime("%m-%d-%Y")
        assignment_data[tech]["workload_history"].append(
            start_date.strftime("%m-%d-%Y")
        )


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

    # Logging to confirm the file creation
    if not os.path.exists(output_file_path):
        logging.warning(
            f"Assignments file not found at {output_file_path}. Creating a new file."
        )
        with open(output_file_path, "w") as new_file:
            new_file.write("Start Date, End Date, Agent 1, Email1, Agent 2, Email2\n")

        logging.info(f"Assignments file created at {output_file_path}")
    else:
        logging.info(f"Assignments written to {output_file_path}")


def generate_rotation_schedule(employee_data, weeks_in_year):
    schedule = {}

    current_date = datetime.now()
    assigned_pairs_queue = initialize_assigned_pairs_queue(4)
    max_assignments = calculate_max_assignments(
        weeks_in_year, len(employee_data[employee_data["Available"] == "yes"])
    )

    # Load assignment history
    history = load_assignment_history()

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

        weights = update_weights(weights, history)  # Pass history to update_weights

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


def main():
    global employee_data, team_list_path, assignments_path

    employee_data = read_employee_data(team_list_path)

    if employee_data is not None:
        # Log backup activity
        log_activity("Backing up existing assignments.")
        backup_existing_assignments()

        # Initialize the assignment data log file
        initialize_assignment_data_log()

        # Generate rotation schedule
        rotation_schedule = generate_rotation_schedule(employee_data, weeks_in_year=52)

        # Log assignment data at the end of script execution
        log_assignment_data(rotation_schedule)  # Pass the assignment data

        # Write the schedule to an Excel file
        write_to_excel(rotation_schedule)

    else:
        logging.error("Exiting program due to errors.")

    # Log the end of the script execution
    log_activity("Script execution completed.\n")


if __name__ == "__main__":
    main()
