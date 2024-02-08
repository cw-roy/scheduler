### Rotation Scheduler

#### Overview

Generates pairings for on-point schedule

### Setup

1. **Python Virtual Environment:**
   - Create a Python virtual environment (venv):
     ```
     python -m venv venv
     ```

2. **Activate Virtual Environment:**
   - On Windows:
     ```
     venv\Scripts\activate.ps1
     ```
   - On macOS/Linux:
     ```
     source venv/bin/activate
     ```

3. **Install Dependencies:**
   - Install required packages from `requirements.txt`:
     ```
     pip install -r requirements.txt
     ```

4. **Run the Script:**
   - Execute the script (`scheduler.py`) within the activated virtual environment.

#### Prerequisites

- Python 3 running in a virtual environment, see setup above
- pandas library installed from requirements.txt or (`pip install pandas`)

#### Usage

1. **Input Data:**
   - Create an Excel spreadsheet (`team_list.xlsx`) containing employee details:
     - <strong><u>Column names must be Name, Email, and Availability (yes/no)</u></strong>
   - Save the file in the "working" directory.

2. **Running the Script:**
   - Execute the script
        - Linux: python3 `scheduler.py`
        - Windows: python `scheduler.py`

3. **Logging:**
   - Logs are recorded in the "logging" directory.
   - `scheduler_events.log` logs script activities.
   - `assignment_data_log.csv` records historical assignment data.

4. **Output:**
   - The rotation schedule is written to `assignments.xlsx` in the "working" directory.
   - Assignments include start/end dates, employee pairs, and email addresses.

5. **Adjustments:**
   - Customize weighting and adjustment logic in the script for fine-tuning.
   - Adjust logging configuration for console/file output.

### Functions

1. **read_employee_data(file_path):**
   - Reads employee data from an Excel spreadsheet.
   - Validates and cleans column names.

2. **check_and_exit_if_no_changes(current_employee_data, previous_employee_data):**
   - Checks for changes in employee availability and exits if none.

3. **initialize_assignment_data_log():**
   - Initializes the assignment data log file if not present.

4. **load_assignment_history():**
   - Loads historical assignment data from `assignment_data_log.csv`.

5. **log_assignment_data(assignment_data):**
   - Logs assignment data to `assignment_data_log.csv`.

6. **log_activity(activity_description):**
   - Logs activities to `scheduler_events.log`.

7. **backup_existing_assignments():**
   - Backs up existing assignment data log.

8. **delete_old_backups(directory, base_name, keep_latest):**
   - Deletes older backup files, keeping the specified number.

9. **get_email_addresses(employee_data, pair):**
   - Gets email addresses for a given pair of employees.

10. **normalize_weights(weights, max_assignments):**
    - Normalizes weights to ensure they add up to `max_assignments`.

11. **detect_changes(old_data, new_data):**
    - Detects changes between two versions of employee data.

12. **calculate_max_assignments(weeks_in_year, available_tech_count):**
    - Calculates the maximum assignments per week.

13. **initialize_assigned_pairs_queue(size):**
    - Initializes the assigned pairs queue with a specified size.

14. **initialize_weights(employee_data):**
    - Initializes weights based on the number of available employees.

15. **load_and_detect_changes(previous_data_path, current_data, log_changes=True):**
    - Loads previous employee data, compares with current data, and logs changes.

16. **calculate_week_dates(current_date, week):**
    - Calculates the start and end dates of a week.

17. **generate_paired_employees(employee_data, weights, assigned_pairs_queue, max_assignments):**
    - Generates a pair of employees for assignment.

18. **update_weights(weights, history=None):**
    - Updates weights based on certain criteria.

19. **adjust_weights(weights, history):**
    - Adjusts weights based on certain criteria and historical data.

20. **update_assignment_data(assignment_data, paired_employees, end_date, start_date):**
    - Updates assignment data for each employee.

21. **write_to_excel(schedule):**
    - Writes the schedule to an Excel file with the specified format.

22. **generate_rotation_schedule(employee_data, weeks_in_year):**
    - Generates the rotation schedule for the entire year.

23. **main():**
    - Main function to execute the script.

#### Notes

- Adjust logging and weighting logic as needed for specific use cases.