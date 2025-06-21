# Import required modules
import os  # For file and directory operations
import re  # For regular expression matching
from datetime import datetime  # For handling date and time
import time  # For implementing timed checks
import sys  # For manipulating sys.path
from io import StringIO  # For capturing stdout
from contextlib import redirect_stdout  # For redirecting stdout
import importlib.util
import yaml

from utility.send_outlook_email import send_outlook_email

# Define user configuration with lists of regular expressions and folder paths
"""
user_config = {
    "folders": [
        {
            "folder_path": r"K:\Chain\2024Q4\Preliminary\Test2 - new",
            "updating_script": r"V:\新增資料夾\updating.py"
        },
        {
            "folder_path": r"K:\Chain\2024Q4\Preliminary\Test2 - interim",
            "updating_script": r"V:\新增資料夾\updating.py"
        },
    ],
    "file_group_a": [r"Data - Section", r"Data - Taxes", r"Data - Ownership", r"Related"],  # List of regex patterns for Group A
    "file_group_b": [r"Data - All", r"Chain Summary", r"BM Compare"],  # List of regex patterns for Group B
    "check_interval": 2,  # Check every 30 seconds
    "cooldown_period": 2  # Cooldown period of 60 seconds
}
"""

with open('config.yaml', 'r', encoding='utf-8') as f:
    user_config = yaml.safe_load(f)



# Define a function to get the last modified time of a file
def get_last_save_time(file_path):
    try:
        last_time = os.path.getmtime(file_path)
        last_time_str = datetime.fromtimestamp(last_time).strftime("%Y-%m-%d %H:%M:%S")
        return last_time, last_time_str
    except FileNotFoundError:
        print_message(f"File {file_path} not found, skipping...", "WARNING")
        return None, None

# Define a function to print formatted messages
def print_message(message, message_type="INFO"):
    symbol = {"INFO": "ℹ", "ACTION": "➤", "WARNING": "⚠", "ERROR": "❌"}
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"{symbol.get(message_type, 'ℹ')} [{message_type}] {current_time}: {message}")

# Define a function to dynamically import and run a script
def run_updating_script(script_path, base_directory):
    script_name = os.path.basename(script_path).replace('.py', '')
    original_argv = sys.argv.copy()
    
    try:
        spec = importlib.util.spec_from_file_location(script_name, script_path)
        if spec is None:
            raise ImportError(f"Cannot load script {script_path}")
        module = importlib.util.module_from_spec(spec)
        sys.modules[script_name] = module

        os.environ["BASE_DIRECTORY_FROM_MONITOR"] = base_directory

        sys.argv = [script_path, base_directory]
        print_message(f"Attempting to run script {script_path} with BASE_DIRECTORY: {base_directory}", "ACTION")
        
        spec.loader.exec_module(module)
        exit_code = module.main()
        
        sys.argv = original_argv
        return exit_code == 0

    except Exception as e:
        print_message(f"Error executing script {script_path}: {str(e)}", "ERROR")
        
        # Dual-Layer Error Notification
        error_subject = f"Critical Error: Updating Script Failed to Execute - {os.path.basename(script_path)}"
        error_body = f"An unhandled error occurred while attempting to execute the updating script '{script_path}' for folder '{base_directory}'.\n\nError details: {str(e)}\n\nPlease investigate immediately."
        
        try:
            send_outlook_email(
                to_recipients=user_config["email_recipients"]["to"],
                subject=error_subject,
                body=error_body,
                cc_recipients=user_config["email_recipients"]["cc"],
                bcc_recipients=user_config["email_recipients"]["bcc"]
            )
            print_message("Critical error notification email sent.", "ACTION")
        except Exception as mail_error:
            print_message(f"Failed to send critical error notification email: {mail_error}", "ERROR")

        return False

    finally:
        if script_name in sys.modules:
            del sys.modules[script_name]
        sys.argv = original_argv
        
          
            
            
# Define the monitoring function for a single folder
def monitor_folder(folder_path):
    group_a_last_times = {}
    group_b_last_times = {}

    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path):
            last_time, last_time_str = get_last_save_time(file_path)
            if last_time is None:
                continue
            for pattern in user_config["file_group_a"]:
                if re.match(pattern, file_name):
                    group_a_last_times[file_name] = last_time
                    break
            for pattern in user_config["file_group_b"]:
                if re.match(pattern, file_name):
                    group_b_last_times[file_name] = last_time
                    break

    if group_a_last_times:
        group_a_newest = max(group_a_last_times.values())
        group_a_newest_str = datetime.fromtimestamp(group_a_newest).strftime("%Y-%m-%d %H:%M:%S")
    else:
        group_a_newest = 0
        group_a_newest_str = "N/A"
    if group_b_last_times:
        group_b_oldest = min(group_b_last_times.values())
        group_b_oldest_str = datetime.fromtimestamp(group_b_oldest).strftime("%Y-%m-%d %H:%M:%S")
    else:
        group_b_oldest = float('inf')
        group_b_oldest_str = "N/A"

    print_message(f"Monitoring {folder_path}, Group A Newest Time: {group_a_newest_str}, Group B Oldest Time: {group_b_oldest_str}")
    return group_a_last_times, group_b_last_times, group_a_newest, group_b_oldest

# Define the monitoring function
def monitor_files():
    folders = user_config["folders"]
    check_interval = user_config["check_interval"]
    cooldown_period = user_config["cooldown_period"]

    # 檢查所有 updating_scripts 是否存在
    for folder in folders:
        script_path = folder["updating_script"]
        if not os.path.exists(script_path):
            print_message(f"Error: Update script {script_path} not found", "ERROR")
            return

    print_message(f"Monitoring started for {len(folders)} folders")

    try:
        while True:
            update_triggered = False
            for folder in folders:
                folder_path = folder["folder_path"]
                updating_script = folder["updating_script"]

                if not os.path.exists(folder_path):
                    print_message(f"Error: Folder path {folder_path} not found", "WARNING")
                    continue

                group_a_last_times, group_b_last_times, group_a_newest, group_b_oldest = monitor_folder(folder_path)

                if group_a_last_times and group_b_last_times:
                    if group_a_newest > group_b_oldest:
                        print_message(
                            f"Group A newer than Group B in {folder_path}, entering cooldown period for {cooldown_period} seconds"
                        )
                        cooldown_start = time.time()
                        while time.time() - cooldown_start < cooldown_period:
                            time.sleep(check_interval)
                            all_stable = True
                            for file_name in group_a_last_times.keys():
                                file_path = os.path.join(folder_path, file_name)
                                if os.path.exists(file_path):
                                    new_time, _ = get_last_save_time(file_path)
                                    if new_time is None:
                                        continue
                                    if new_time > group_a_last_times[file_name]:
                                        all_stable = False
                                        group_a_last_times[file_name] = new_time
                                        print_message(
                                            f"Group A file {file_name} in {folder_path} updated, extending cooldown"
                                        )
                                        cooldown_start = time.time()
                                        break
                            if all_stable:
                                print_message(f"Cooldown ended for {folder_path}, executing update script")
                                if run_updating_script(updating_script, folder_path):
                                    print_message(f"Update script {updating_script} executed successfully")
                                    update_triggered = True
                                else:
                                    print_message(f"Update script {updating_script} failed to execute", "WARNING")
                                break
                    else:
                        print_message(f"Group A not newer than Group B in {folder_path}, moving to next folder")
                else:
                    print_message(f"Not all group files found in {folder_path}, moving to next folder")

                if update_triggered:
                    print_message("Update triggered, breaking folder loop", "ACTION")
                    break

            if not update_triggered:
                print_message(f"Waiting {check_interval} seconds for the next check")
                time.sleep(check_interval)

    except KeyboardInterrupt:
        print_message("Monitoring stopped manually", "WARNING")
       
        
if __name__ == "__main__":
    monitor_files()

# log_filepath
