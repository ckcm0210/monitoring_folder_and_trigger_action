# File Monitoring and Automated Update System

This is an automated Python solution designed to monitor file changes in specific folders. When it detects that a predefined set of files (Group A) is newer than another set of files (Group B), it will automatically trigger an update script and send an email notification upon completion of the update.

## Project Overview

This project primarily consists of the following modules:

* **`main_workflow.py`**: The entry point for the entire automation process. It is responsible for initializing the file monitoring system and coordinating the execution of various modules during the monitoring process.
* **`monitoring.py`**: The core file monitoring module. It periodically checks the folders defined in the configuration, comparing the last modification times of files in Group A and Group B to determine if an update needs to be triggered.
* **`updating.py`**: When `monitoring.py` detects file changes and meets the triggering conditions, this script is called to perform the actual update operation. This script supports operations on Excel files (e.g., executing macros, handling passwords) and generates operation logs.
* **`send_outlook_email.py`**: A utility module used to send email notifications via Outlook. It is used by `updating.py` to send notification emails containing log content after the update is complete.
* **`monitoring_config.yaml`**: The configuration file for the monitoring module, defining the folders to be monitored, regular expression patterns for file groups, check intervals, and cooldown periods.
* **`updating_config.yaml`**: The configuration file for the updating module, defining the log directory, macro and password settings for Excel files, email recipients, and subject prefixes.

## File Structure

To ensure that the program can correctly locate and import all modules and configuration files, your project directory structure should ideally be as follows:

your_project_root/<br/>
├── main_workflow.py<br/>
├── monitoring.py<br/>
├── monitoring_config.yaml<br/>
├── updating.py<br/>
├── updating_config.yaml<br/>
└── utility/<br/>
     └── send_outlook_email.py<br/>

* `your_project_root/`: This is your main project directory.
* `main_workflow.py`, `monitoring.py`, `updating.py`: These are the main Python scripts, located directly under the project root.
* `monitoring_config.yaml`, `updating_config.yaml`: These are the configuration files, also located directly under the project root.
* `utility/`: This is a subdirectory used to store auxiliary Python modules.
* `send_outlook_email.py`: This module is located within the `utility/` subdirectory and is imported and used by other main scripts.

This structure adheres to Python's modular best practices and facilitates management.

## System Features

* **Continuous File Monitoring**: Periodically scans predefined folders.
* **Intelligent Change Detection**: Determines whether to trigger an update based on the modification times of Group A and Group B files.
* **Stability Cooldown Period**: After file changes are detected, waits for a cooldown period to ensure files are stable before executing the update script, preventing triggers while files are still being written.
* **Automated Update Trigger**: Automatically executes the predefined update script when conditions are met.
* **Excel Automation**: `updating.py` can process Excel files, including opening password-protected files, executing VBA macros, and saving and closing files.
* **Detailed Logging**: Records all important operations and potential issues.
* **Email Notification**: Sends detailed operation logs as email content via Outlook after the update process is complete.
* **Configurability**: Most key settings are stored in `.yaml` configuration files, allowing users to easily modify them as needed.

## Environment Requirements

1.  **Python Environment**: Ensure Python 3.x is installed on your system.
2.  **Required Libraries**:
    * **Recommended Installation**: Run `pip install -r requirements.txt` in the project root directory.
    * Manual Installation (if no `requirements.txt`):
        * `pywin32` (`pip install pywin32`): Used for interacting with Outlook and Excel applications.
        * `openpyxl` (`pip install openpyxl`): Used for processing Excel XLSX/XLSM files.
        * `PyYAML` (`pip install PyYAML`): Used for reading `config.yaml` files.
3.  **Outlook Application**: The `send_outlook_email.py` script relies on the Outlook application being open and running.

## Installation

1.  **Clone the Repository**:
    ```bash
    git clone <Your_GitHub_Repository_URL>
    cd <Your_Project_Folder_Name>
    ```

2.  **Install Dependencies**:
    It is recommended to use `pip` to install all necessary Python libraries.
    ```bash
    pip install pywin32 pyyaml openpyxl
    ```
    * `pywin32`: Provides `win32com.client` for interacting with Windows applications (like Outlook and Excel).
    * `pyyaml`: Used for reading and parsing `.yaml` configuration files.
    * `openpyxl`: Used for handling `.xlsx` format Excel files (may be used in `updating.py`, although the primary dependency is `win32com`).

## Configuration

Before running the program, you need to configure the following two `.yaml` files to suit your environment and requirements:

1.  **`monitoring_config.yaml`**:
    * `folders`: Defines the paths of the folders to be monitored and their corresponding update script paths.
    * `file_group_a`: A list of regular expression patterns to match files for monitoring (typically data input or update files).
    * `file_group_b`: Another list of regular expression patterns to match files used as a comparison baseline (typically output or summary files).
    * `check_interval`: The waiting time (in seconds) between each check for file changes.
    * `cooldown_period`: The cooldown time (in seconds) before executing the update script after files are detected as stable.

    Example:
    ```yaml
    folders:
      - folder_path: "K:\\Chain\\2024Q4\\Preliminary\\Test2 - new"
        updating_script: "V:\\新增資料夾\\updating.py" # Note: The updating_script path here still points to the actual updating.py file
      - folder_path: "K:\\Chain\\2024Q4\\Preliminary\\Test2 - interim"
        updating_script: "V:\\新增資料夾\\updating.py" # Same as above
    file_group_a:
      - "Data - Section"
      - "Data - Taxes"
      - "Data - Ownership"
      - "Related"
    # ... Other configurations
    ```

2.  **`updating_config.yaml`**:
    * `email_recipients`: Configures the `to`, `cc`, and `bcc` recipient lists for emails.
    * `email_subject_prefix`: The prefix for email subjects.
    * `log_directory`: The directory where program log files will be stored.
    * `file_configs`: Configurations specific to certain Excel files, such as macro names, open passwords, and write passwords.
    * `advanced_settings`: Advanced settings, such as whether the Excel application is visible, number of retries, etc.

    Example:
    ```yaml
    email_recipients:
      to: ["your_email@example.com"]
      cc: []
      bcc: []
    file_configs:
      Data - All:
        macro: null
        open_password: null
        write_password: "2011chainref"
      Chain Summary:
        macro: null
        open_password: null
        write_password: "2011chainref"
      BM Compare:
        macro: "Main"
        open_password: null
        write_password: null
    # ... Other configurations
    ```
    **Note**: The `updating_script` path specified in `monitoring_config.yaml` should point to the actual `updating.py` file.

## How to Run

Ensure that you have completed the installation and configuration steps mentioned above.

1.  **Ensure Outlook is Running**: Since the email functionality relies on Outlook, please ensure your Outlook application is running.
2.  **Run the Main Workflow**:
    Open your command prompt or terminal, navigate to the project's root directory, and run `main_workflow.py`:
    ```bash
    python main_workflow.py
    ```

    Alternatively, if you are in a Jupyter Notebook environment, you can directly run the `main_workflow_entry()` function:
    ```python
    # In Jupyter Notebook
    import sys
    import os

    current_dir = os.path.dirname(os.path.abspath(__file__))
    # Add the project root directory to sys.path to ensure the utility module can be found
    project_root = os.path.abspath(os.path.join(current_dir)) # Ensure project_root is the directory of the current file
    if project_root not in sys.path:
        sys.path.append(project_root)

    import main_workflow
    main_workflow.main_workflow_entry()
    ```

The program will start monitoring files. You can press `Ctrl+C` at any time to manually stop the monitoring process.

## Important Notes

* **File Paths**: Please ensure all file and folder paths in the configuration are correct and accessible.
* **Outlook and Excel Permissions**: Ensure that the user account running the script has sufficient permissions to control Outlook and Excel applications.
* **Log Directory**: Ensure that the `log_directory` configured in `updating_config.yaml` exists and is writable.
* **Excel Passwords**: If Excel files are password-protected, please configure `open_password` or `write_password` correctly in `updating_config.yaml`.
* **Macro Security**: If your Excel files contain macros, please ensure Excel's macro security settings allow macros to run, or add your project folder to trusted locations.
* **Filename Spaces**: Please ensure all imported Python filenames do not contain spaces, e.g., `send_outlook_email.py` instead of `send_outlook_email .py`.

## Troubleshooting

* **`ImportError`**: If you encounter an `ImportError` during runtime, please check if the file names are correct and ensure all necessary modules are installed. Specifically, check if the filename for `send_outlook_email.py` has been corrected (from `send_outlook_email .py` to `send_outlook_email.py`) and its path conforms to `utility/send_outlook_email.py`.
* **Outlook/Excel Automation Issues**:
    * Ensure Outlook and Excel are installed and running correctly.
    * Check for any security pop-ups that might be blocking program interaction.
    * Check if `pywin32` is installed correctly.
* **Files Not Triggering Update**:
    * Check if the `file_group_a` and `file_group_b` regular expressions in `monitoring_config.yaml` correctly match your files.
    * Check if the modification times of the files are as expected, i.e., Group A is indeed newer than Group B.
    * Adjust `check_interval` and `cooldown_period` to suit your needs.

## Contribution

If you have any suggestions or find errors, feel free to open an issue or submit a pull request.

## License

[Place your license information here, e.g., MIT License]
