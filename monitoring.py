import os
import re
from datetime import datetime
import time
import sys
import importlib.util
import yaml

from utility.send_outlook_email import send_outlook_email

# --- Config 讀取與全域變數 ---
def load_monitoring_config(path):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        return config
    except Exception as e:
        print(f"[ERROR] Cannot load config: {e}")
        sys.exit(1)

monitoring_config = load_monitoring_config('monitoring_config.yaml')

# Helper：expandvars 路徑
def expand_path(path):
    return os.path.expandvars(path)

# --- 工具函式 ---
def print_message(message, message_type="INFO"):
    symbol = {"INFO": "ℹ", "ACTION": "➤", "WARNING": "⚠", "ERROR": "❌"}
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"{symbol.get(message_type, 'ℹ')} [{message_type}] {current_time}: {message}")

def get_last_save_time(file_path):
    try:
        last_time = os.path.getmtime(file_path)
        last_time_str = datetime.fromtimestamp(last_time).strftime("%Y-%m-%d %H:%M:%S")
        return last_time, last_time_str
    except FileNotFoundError:
        print_message(f"File {file_path} not found, skipping...", "WARNING")
        return None, None

def run_updating_script(updating_script_path, monitored_folder_path):
    script_name = os.path.basename(updating_script_path).replace('.py', '')
    original_argv = sys.argv.copy()
    try:
        spec = importlib.util.spec_from_file_location(script_name, updating_script_path)
        if spec is None:
            raise ImportError(f"Cannot load script {updating_script_path}")
        module = importlib.util.module_from_spec(spec)
        sys.modules[script_name] = module

        os.environ["BASE_DIRECTORY_FROM_MONITOR"] = monitored_folder_path
        sys.argv = [updating_script_path, monitored_folder_path]
        print_message(f"Attempting to run script {updating_script_path} with BASE_DIRECTORY: {monitored_folder_path}", "ACTION")
        spec.loader.exec_module(module)
        exit_code = module.main()
        sys.argv = original_argv
        return exit_code == 0
    except Exception as e:
        print_message(f"Error executing script {updating_script_path}: {str(e)}", "ERROR")
        # Optional: email notification if config 裡有 email_recipients
        if "email_recipients" in monitoring_config:
            try:
                send_outlook_email(
                    to_recipients=monitoring_config["email_recipients"].get("to", []),
                    subject=f"Critical Error: Updating Script Failed - {os.path.basename(updating_script_path)}",
                    body=f"Updating script '{updating_script_path}' for folder '{monitored_folder_path}' failed.\n\nError details: {str(e)}",
                    cc_recipients=monitoring_config["email_recipients"].get("cc", []),
                    bcc_recipients=monitoring_config["email_recipients"].get("bcc", [])
                )
                print_message("Critical error notification email sent.", "ACTION")
            except Exception as mail_error:
                print_message(f"Failed to send critical error notification email: {mail_error}", "ERROR")
        return False
    finally:
        if script_name in sys.modules:
            del sys.modules[script_name]
        sys.argv = original_argv

def monitor_folder(monitored_folder_path):
    group_a_last_times = {}
    group_b_last_times = {}
    file_group_a = monitoring_config.get("file_group_a", [])
    file_group_b = monitoring_config.get("file_group_b", [])

    for file_name in os.listdir(monitored_folder_path):
        file_path = os.path.join(monitored_folder_path, file_name)
        if os.path.isfile(file_path):
            last_time, _ = get_last_save_time(file_path)
            if last_time is None:
                continue
            # 支援 regex 或 substring
            for pattern in file_group_a:
                if re.search(pattern, file_name):
                    group_a_last_times[file_name] = last_time
                    break
            for pattern in file_group_b:
                if re.search(pattern, file_name):
                    group_b_last_times[file_name] = last_time
                    break

    group_a_newest = max(group_a_last_times.values()) if group_a_last_times else 0
    group_b_oldest = min(group_b_last_times.values()) if group_b_last_times else float('inf')

    group_a_newest_str = datetime.fromtimestamp(group_a_newest).strftime("%Y-%m-%d %H:%M:%S") if group_a_last_times else "N/A"
    group_b_oldest_str = datetime.fromtimestamp(group_b_oldest).strftime("%Y-%m-%d %H:%M:%S") if group_b_last_times else "N/A"

    print_message(f"Monitoring {monitored_folder_path}, Group A Newest Time: {group_a_newest_str}, Group B Oldest Time: {group_b_oldest_str}")
    return group_a_last_times, group_b_last_times, group_a_newest, group_b_oldest

def monitor_files():
    folders = monitoring_config.get("folders", [])
    check_interval = monitoring_config.get("check_interval", 2)
    cooldown_period = monitoring_config.get("cooldown_period", 2)

    # 檢查所有 updating_scripts 是否存在
    for folder in folders:
        updating_script_path = expand_path(folder["updating_script"])
        if not os.path.exists(updating_script_path):
            print_message(f"Error: Update script {updating_script_path} not found", "ERROR")
            return

    print_message(f"Monitoring started for {len(folders)} folders")

    try:
        while True:
            update_triggered = False
            for folder in folders:
                monitored_folder_path = expand_path(folder["folder_path"])
                updating_script_path = expand_path(folder["updating_script"])
                if not os.path.exists(monitored_folder_path):
                    print_message(f"Error: Folder path {monitored_folder_path} not found", "WARNING")
                    continue
                group_a_last_times, group_b_last_times, group_a_newest, group_b_oldest = monitor_folder(monitored_folder_path)
                if group_a_last_times and group_b_last_times:
                    if group_a_newest > group_b_oldest:
                        print_message(f"Group A newer than Group B in {monitored_folder_path}, entering cooldown period for {cooldown_period} seconds")
                        cooldown_start = time.time()
                        while time.time() - cooldown_start < cooldown_period:
                            time.sleep(check_interval)
                            all_stable = True
                            for file_name in group_a_last_times.keys():
                                file_path = os.path.join(monitored_folder_path, file_name)
                                if os.path.exists(file_path):
                                    new_time, _ = get_last_save_time(file_path)
                                    if new_time is None:
                                        continue
                                    if new_time > group_a_last_times[file_name]:
                                        all_stable = False
                                        group_a_last_times[file_name] = new_time
                                        print_message(f"Group A file {file_name} in {monitored_folder_path} updated, extending cooldown")
                                        cooldown_start = time.time()
                                        break
                            if all_stable:
                                print_message(f"Cooldown ended for {monitored_folder_path}, executing update script")
                                if run_updating_script(updating_script_path, monitored_folder_path):
                                    print_message(f"Update script {updating_script_path} executed successfully")
                                    update_triggered = True
                                else:
                                    print_message(f"Update script {updating_script_path} failed to execute", "WARNING")
                                break
                    else:
                        print_message(f"Group A not newer than Group B in {monitored_folder_path}, moving to next folder")
                else:
                    print_message(f"Not all group files found in {monitored_folder_path}, moving to next folder")
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
