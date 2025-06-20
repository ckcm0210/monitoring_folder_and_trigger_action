# Automated File Update and Monitoring System

## Project Overview

This project is designed to automatically monitor changes in files within specified folders and, when certain conditions are met, trigger a Python script (`updating.py`) to perform file update operations. The system incorporates a dual-layer error notification mechanism to ensure timely alerts in case of any anomalies.

### Prerequisites

1.  **Python Environment**: Ensure Python 3.x is installed on your system.
2.  **Required Libraries**:
    * **Recommended Installation**: In the project root directory, run `pip install -r requirements.txt`.
    * Manual Installation (if `requirements.txt` is not used):
        * `pywin32` (`pip install pywin32`): For interacting with the Outlook application.
        * `openpyxl` (`pip install openpyxl`): For handling Excel XLSX/XLSM files.
        * `PyYAML` (`pip install PyYAML`): For reading the `config.yaml` file.
3.  **Outlook Application**: The `send_outlook_email.py` script relies on the Outlook application running.

### Project Files

* **`monitoring_folder_and_trigger_action.py`**:
    The main monitoring script responsible for listening to file changes in folders, determining if trigger conditions are met, and executing `updating.py`.
* **`updating.py`**:
    The script that executes the actual file update logic, including Excel file processing, macro execution, and email notifications.
* **`send_outlook_email.py`**:
    A general utility script used for sending Outlook emails.
* **`config.yaml`**:
    The externalized configuration file containing all configurable paths, file rules, email lists, etc.

## Configuration

All project-related settings are centralized in the `config.yaml` file. It is crucial to review and modify this file to match your environment before starting the project.

### `config.yaml` Example and Explanation

```yaml
# This section defines the folders to be monitored.
# Each folder has a path and an associated updating script.
folders:
  - folder_path: "K:\\Chain\\2024Q4\\Preliminary\\Test2 - new"
    updating_script: "V:\\新增資料夾\\updating.py"
  - folder_path: "K:\\Chain\\2024Q4\\Preliminary\\Test2 - interim"
    updating_script: "V:\\新增資料夾\\updating.py"

# These are regex patterns for files in Group A.
file_group_a:
  - "Data - Section"
  - "Data - Taxes"
  - "Data - Ownership"
  - "Related"

# These are regex patterns for files in Group B.
file_group_b:
  - "Data - All"
  - "Chain Summary"
  - "BM Compare"

# How often to check for file changes (in seconds).
check_interval: 2

# Cooldown period after files are detected as stable (in seconds).
cooldown_period: 2

# Email recipient settings for notifications.
email_recipients:
  to: ["your_email@example.com"]
  cc: []
  bcc: []

# Prefix for email subjects.
email_subject_prefix: "K Chain File Refresh Completion Notification"

# Directory where log files will be stored.
log_directory: "D:\\Pzone\\log"
