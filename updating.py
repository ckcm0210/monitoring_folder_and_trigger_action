import os
import win32com.client as win32
import time
import logging
from openpyxl import load_workbook
from datetime import datetime
from pathlib import Path
from utility.send_outlook_email import send_outlook_email
import yaml

# --- Config 讀取與全域變數 ---
def load_updating_config(path):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        return config
    except Exception as e:
        print(f"[ERROR] Cannot load config: {e}")
        exit(1)

updating_config = load_updating_config('updating_config.yaml')

# 支援外部 BASE_DIRECTORY 覆蓋
if os.environ.get("BASE_DIRECTORY_FROM_MONITOR"):
    updating_config["base_directory"] = os.environ["BASE_DIRECTORY_FROM_MONITOR"]

to_recipients = updating_config["email_recipients"]["to"]
cc_recipients = updating_config["email_recipients"]["cc"]
bcc_recipients = updating_config["email_recipients"]["bcc"]
email_subject = updating_config["email_subject_prefix"]
log_directory = updating_config["log_directory"]
file_configs = updating_config["file_configs"]
advanced_settings = updating_config["advanced_settings"]
base_directory = updating_config["base_directory"]

logger = None

class ExcelAutomationError(Exception):
    pass

def setup_logging():
    global logger
    Path(log_directory).mkdir(parents=True, exist_ok=True)
    current_time = datetime.now()
    log_filename = f"ExcelAutoRefresh_{current_time.strftime('%Y%m%d_%H%M%S')}.log"
    log_filepath = os.path.join(log_directory, log_filename)
    os.environ["log_filepath"] = log_filepath

    logger = logging.getLogger('ExcelAutomation')
    logger.handlers.clear()
    file_handler = logging.FileHandler(log_filepath, encoding='utf-8')
    file_formatter = logging.Formatter('%(asctime)s | %(levelname)-8s | %(message)s')
    file_handler.setFormatter(file_formatter)
    console_handler = logging.StreamHandler()
    console_formatter = logging.Formatter('[%(asctime)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    console_handler.setFormatter(console_formatter)
    logger.setLevel(logging.INFO)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    logger.propagate = False

    logger.info("=" * 80)
    logger.info("🚀 Excel automation program started")
    logger.info(f"📁 Log directory created/confirmed: {log_directory}")
    logger.info(f"📄 Log file: {log_filename}")
    logger.info(f"📁 Configured base directory: {base_directory}")
    logger.info(f"🏷️ Configured file prefixes: {list(file_configs.keys())}")
    logger.info("=" * 80)
    return logger

def console_print(message, level='info'):
    if logger is None:
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"[{timestamp}] {message}")
        return
    if message == "":
        message = " "
    if level.lower() == 'info':
        logger.info(message)
    elif level.lower() == 'warning':
        logger.warning(message)
    elif level.lower() == 'error':
        logger.error(message)
    else:
        logger.info(message)

def safe_execute(func, *args, **kwargs):
    max_retries = advanced_settings["max_retries"]
    retry_delay_base = advanced_settings["retry_delay_base"]
    for attempt in range(max_retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            if attempt < max_retries - 1:
                console_print(f"Operation failed, retrying ({attempt + 1}/{max_retries}): {str(e)}", level='warning')
                time.sleep(retry_delay_base ** attempt)
            else:
                raise e

def get_file_last_save_time(file_path):
    if not os.path.exists(file_path):
        return None
    try:
        timestamp = os.path.getmtime(file_path)
        last_save_time = datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')
        return last_save_time
    except Exception as e:
        console_print(f"Cannot get last save time for file '{file_path}': {str(e)}", level='warning')
        return None

def is_excel_file_accessible(file_path, open_password=None):
    if open_password is not None:
        console_print(f"File has password protection, skipping accessibility check")
        return True
    try:
        wb = load_workbook(file_path, read_only=True, data_only=True)
        wb.close()
        return True
    except Exception as e:
        console_print(f"File '{file_path}' is not accessible: {str(e)}", level='error')
        return False

def get_last_save_author_improved(file_path, has_password=False, open_password=None, workbook_obj=None):
    # Method 1: If workbook object is provided, get author from opened workbook
    if workbook_obj is not None:
        try:
            builtin_props = workbook_obj.BuiltinDocumentProperties
            last_author = builtin_props("Last Author").Value
            return last_author if last_author else "Last author info not found from opened workbook"
        except Exception as e:
            console_print(f"Cannot get author info from opened workbook: {str(e)}", level='warning')
    # Method 2: For password-protected files, use win32com to open directly and extract
    if has_password and open_password is not None:
        excel_app = None
        workbook = None
        try:
            console_print(f"Using win32com to open password-protected file for metadata extraction...")
            excel_app = win32.Dispatch("Excel.Application")
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            excel_app.EnableEvents = False
            open_params = {
                'Filename': file_path,
                'Password': open_password,
                'ReadOnly': True,
                'UpdateLinks': False,
                'IgnoreReadOnlyRecommended': True
            }
            workbook = excel_app.Workbooks.Open(**open_params)
            builtin_props = workbook.BuiltinDocumentProperties
            last_author = builtin_props("Last Author").Value
            return last_author if last_author else "Last author info not found in password-protected file"
        except Exception as e:
            console_print(f"Failed to open password file with win32com: {str(e)}", level='warning')
            return f"Password-protected file, cannot extract author info: {str(e)}"
        finally:
            if workbook:
                try:
                    workbook.Close(SaveChanges=False)
                except:
                    pass
            if excel_app:
                try:
                    excel_app.Quit()
                except:
                    pass
    # Method 3: Use openpyxl for files without password
    if not has_password:
        try:
            wb = load_workbook(file_path, read_only=True, data_only=True)
            last_author = wb.properties.lastModifiedBy
            wb.close()
            return last_author if last_author else "Last author info not found from openpyxl"
        except Exception as e:
            console_print(f"Cannot get author info via openpyxl: {str(e)}", level='warning')
    # Default
    if has_password:
        return "Password-protected file, openpyxl cannot get author info"
    else:
        return "Cannot get author info"

def get_workbook_metadata_via_win32com(file_path, open_password=None, write_password=None):
    excel_app = None
    workbook = None
    metadata = {}
    try:
        console_print(f"Using win32com to extract file metadata...")
        excel_app = win32.Dispatch("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_app.EnableEvents = False
        open_params = {
            'Filename': file_path,
            'ReadOnly': True,
            'UpdateLinks': False,
            'IgnoreReadOnlyRecommended': True
        }
        if open_password:
            open_params['Password'] = open_password
        if write_password:
            open_params['WriteResPassword'] = write_password
        workbook = excel_app.Workbooks.Open(**open_params)
        builtin_props = workbook.BuiltinDocumentProperties
        try:
            last_author = builtin_props("Last Author").Value
            metadata["👤 Last Author"] = str(last_author) if last_author is not None else "Not set"
        except Exception:
            metadata["👤 Last Author"] = "Unable to retrieve"
        try:
            last_save_time = builtin_props("Last Save Time").Value
            if last_save_time is not None and isinstance(last_save_time, datetime):
                last_save_time = last_save_time.strftime('%Y-%m-%d %H:%M:%S')
            metadata["🕒 Last Save Time"] = str(last_save_time) if last_save_time is not None else "Not set"
        except Exception:
            metadata["🕒 Last Save Time"] = "Unable to retrieve"
        console_print(f"Successfully extracted metadata:")
        console_print(f"   👤 Last author: {metadata.get('👤 Last Author', 'N/A')}")
        console_print(f"   🕒 Last save time: {metadata.get('🕒 Last Save Time', 'N/A')}")
        return metadata
    except Exception as e:
        console_print(f"Failed to extract metadata using win32com: {str(e)}", level='error')
        return {
            "👤 Last Author": f"Extraction failed: {str(e)}",
            "🕒 Last Save Time": f"Extraction failed: {str(e)}"
        }
    finally:
        if workbook:
            try:
                workbook.Close(SaveChanges=False)
            except Exception as e:
                console_print(f"Error occurred while closing workbook: {str(e)}", level='warning')
        if excel_app:
            try:
                excel_app.EnableEvents = True
                excel_app.Quit()
            except Exception as e:
                console_print(f"Error occurred while closing Excel application: {str(e)}", level='warning')

def refresh_workbook_connections(workbook):
    refresh_count = 0
    try:
        excel_links = workbook.LinkSources(Type=win32.constants.xlExcelLinks)
        if excel_links:
            console_print(f"📊 Found {len(excel_links)} Excel file links")
            for i, link_path in enumerate(excel_links):
                link_file_name = os.path.basename(link_path)
                link_dir = os.path.dirname(link_path)
                console_print(f"  └─ Link {i+1}: {link_file_name}")
                console_print(f"     📁 Path: {link_dir}")
                if os.path.exists(link_path):
                    link_author = get_last_save_author_improved(link_path, False)
                    link_last_save_time = get_file_last_save_time(link_path)
                    console_print(f"     👤 Last author: {link_author}")
                    console_print(f"     🕒 Last save time: {link_last_save_time}")
                    try:
                        console_print(f"     🔄 Updating link...")
                        workbook.UpdateLink(Name=link_path, Type=win32.constants.xlExcelLinks)
                        console_print(f"     ✅ Link update successful: {link_file_name}")
                        refresh_count += 1
                    except Exception as e:
                        console_print(f"     ❌ Link update failed: {str(e)}", level='error')
                        refresh_count += 1
                else:
                    console_print(f"     ⚠️ Linked file does not exist: {link_path}", level='warning')
                    console_print(f"     ⏭️ Skipping this link update")
            console_print("")
        connections = workbook.Connections
        if connections and connections.Count > 0:
            console_print(f"🔗 Found {connections.Count} data connections")
            for i in range(1, connections.Count + 1):
                try:
                    connection = connections.Item(i)
                    connection_name = connection.Name
                    console_print(f"  └─ Connection {i}: {connection_name}")
                    console_print(f"     🔄 Refreshing connection...")
                    safe_execute(connection.Refresh)
                    time.sleep(1)
                    refresh_count += 1
                    console_print(f"     ✅ Refresh completed")
                except Exception as e:
                    console_print(f"     ❌ Refresh failed: {str(e)}", level='error')
            console_print("")
        if advanced_settings["force_calculation"]:
            console_print("🧮 Performing full formula recalculation...")
            workbook.Application.CalculateFullRebuild()
            console_print("   ✅ Formula calculation completed")
            console_print("")
    except Exception as e:
        console_print(f"Error occurred while refreshing links: {str(e)}", level='error')
        raise ExcelAutomationError(f"Cannot refresh workbook links: {str(e)}")
    return refresh_count

def execute_macro_safely(excel_app, macro_name):
    try:
        console_print(f"⚙️ Attempting to execute macro: {macro_name}")
        safe_execute(excel_app.Run, macro_name)
        console_print(f"   ✅ Macro execution successful")
        return True
    except Exception as e:
        console_print(f"   ❌ Macro execution failed: {str(e)}", level='error')
        return False

def automate_excel_refresh_links(excel_file_path, file_config):
    macro_to_run = file_config.get("macro")
    file_open_password = file_config.get("open_password")
    file_write_password = file_config.get("write_password")
    file_name = os.path.basename(excel_file_path)
    console_print("")
    console_print(f"📁 Processing file: {file_name}")
    console_print("=" * 60)
    if not os.path.exists(excel_file_path):
        console_print(f"❌ File does not exist: {file_name}", level='error')
        return False
    if not is_excel_file_accessible(excel_file_path, file_open_password):
        console_print(f"❌ File is not accessible, skipping processing: {file_name}", level='error')
        return False
    excel_app = None
    workbook = None
    success = False
    has_password = file_open_password is not None
    metadata_before = {}
    metadata_after = {}
    try:
        if has_password:
            console_print("📋 Getting metadata for password-protected file before processing...")
            metadata_before = get_workbook_metadata_via_win32com(
                excel_file_path,
                file_open_password,
                file_write_password
            )
        else:
            console_print("📋 Getting file metadata before processing...")
            last_save_time_before = get_file_last_save_time(excel_file_path)
            last_author_before = get_last_save_author_improved(
                excel_file_path,
                has_password,
                file_open_password
            )
            metadata_before = {
                "👤 Last Author": last_author_before,
                "🕒 Last Save Time": last_save_time_before
            }
            console_print(f"   🕒 Last save time before processing: {last_save_time_before}")
            console_print(f"   👤 Last author before processing: {last_author_before}")
        console_print("")
        console_print("🚀 Starting Excel application for processing...")
        excel_app = win32.Dispatch("Excel.Application")
        excel_app.Visible = advanced_settings["excel_visible"]
        excel_app.DisplayAlerts = False
        excel_app.EnableEvents = False
        console_print("   ✅ Excel application startup completed")
        console_print(f"📂 Opening file for processing: {file_name}")
        open_params = {
            'Filename': excel_file_path,
            'UpdateLinks': 3,
            'ReadOnly': False,
            'IgnoreReadOnlyRecommended': True,
            'Origin': win32.constants.xlWindows
        }
        if file_open_password:
            open_params['Password'] = file_open_password
        if file_write_password:
            open_params['WriteResPassword'] = file_write_password
        workbook = safe_execute(excel_app.Workbooks.Open, **open_params)
        console_print(f"   ✅ File opened successfully for processing")
        console_print("")
        console_print("🔄 Starting to refresh all external links and data connections...")
        refresh_count = refresh_workbook_connections(workbook)
        console_print(f"✅ Total refreshed {refresh_count} links/connections")
        console_print("")
        if macro_to_run:
            execute_macro_safely(excel_app, macro_to_run)
        else:
            console_print("ℹ️ No macro specified to execute")
        console_print("")
        console_print(f"💾 Saving file: {file_name}")
        console_print(f"   📊 Workbook modification status: {'Modified' if not workbook.Saved else 'Not modified'}")
        console_print(f"   🔒 Workbook read/write status: {'Read-only' if workbook.ReadOnly else 'Writable'}")
        safe_execute(workbook.Save)
        console_print(f"   ✅ File saved successfully")
        console_print("")
        console_print("📋 Getting final file metadata after processing...")
        try:
            builtin_props = workbook.BuiltinDocumentProperties
            try:
                last_author_after = builtin_props("Last Author").Value
                metadata_after["👤 Last Author"] = str(last_author_after) if last_author_after is not None else "Not set"
            except Exception:
                metadata_after["👤 Last Author"] = "Unable to retrieve"
            last_save_time_after = get_file_last_save_time(excel_file_path)
            metadata_after["🕒 Last Save Time"] = last_save_time_after
        except Exception as e:
            console_print(f"Error occurred while getting metadata after processing: {str(e)}", level='warning')
            metadata_after = {
                "👤 Last Author": "Retrieval failed",
                "🕒 Last Save Time": get_file_last_save_time(excel_file_path)
            }
        console_print(f"   👤 Last author after processing: {metadata_after.get('👤 Last Author')}")
        console_print(f"   🕒 Last save time after processing: {metadata_after.get('🕒 Last Save Time')}")
        console_print("📊 Metadata comparison:")
        author_before = metadata_before.get("👤 Last Author")
        author_after = metadata_after.get("👤 Last Author")
        if author_before and author_after:
            if author_before != author_after:
                console_print(f"   👤 Author changed: {author_before} → {author_after}")
            else:
                console_print(f"   👤 Author unchanged: {author_after}")
        else:
            console_print(f"   👤 Author after processing: {author_after}")
        time_before = metadata_before.get("🕒 Last Save Time")
        time_after = metadata_after.get("🕒 Last Save Time")
        if time_before and time_after:
            if time_before != time_after:
                console_print(f"   🕒 Save time updated: {time_before} → {time_after}")
            else:
                console_print(f"   🕒 Save time unchanged: {time_after}")
        else:
            console_print(f"   🕒 Save time after processing: {time_after}")
        success = True
    except Exception as e:
        console_print(f"❌ Error occurred while processing file: {str(e)}", level='error')
        success = False
    finally:
        if workbook:
            try:
                workbook.Close(SaveChanges=False)
                console_print("🔐 Workbook closed")
            except Exception as e:
                console_print(f"⚠️ Error occurred while closing workbook: {str(e)}", level='warning')
        if excel_app:
            try:
                excel_app.EnableEvents = True
                excel_app.Quit()
                console_print("🔐 Excel application closed")
            except Exception as e:
                console_print(f"⚠️ Error occurred while closing Excel application: {str(e)}", level='warning')
    status_icon = "✅" if success else "❌"
    console_print(f"{status_icon} File '{file_name}' processing {'successful' if success else 'failed'}")
    console_print("=" * 60)
    return success

def process_excel_files_in_directory(base_directory, file_configs):
    console_print("")
    console_print(f"🚀 Starting batch processing directory: {base_directory}")
    console_print(f"⏰ Processing start time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    console_print("=" * 70)
    if not os.path.isdir(base_directory):
        console_print(f"❌ Directory does not exist: {base_directory}", level='error')
        return
    processed_files = []
    failed_files = []
    skipped_files = []
    all_excel_files = [f for f in os.listdir(base_directory)
                      if f.lower().endswith(('.xlsx', '.xlsm')) and
                      os.path.isfile(os.path.join(base_directory, f))]
    console_print(f"📊 Found {len(all_excel_files)} Excel files in directory")
    console_print("")
    for prefix in file_configs.keys():
        console_print(f"🔍 Searching for files with prefix: {prefix}")
        matched_files = [f for f in all_excel_files if f.startswith(prefix)]
        if not matched_files:
            console_print(f"⚠️ No files found with prefix: {prefix}", level='warning')
            skipped_files.append(f"No files for prefix: {prefix}")
            continue
        for filename in matched_files:
            full_file_path = os.path.join(base_directory, filename)
            console_print(f"🎯 Found matching file: {filename} (prefix: '{prefix}')")
            current_file_config = file_configs[prefix]
            if automate_excel_refresh_links(full_file_path, current_file_config):
                processed_files.append(filename)
            else:
                failed_files.append(filename)
    for filename in all_excel_files:
        if not any(filename in processed_files or filename in failed_files for filename in all_excel_files):
            console_print(f"⏭️ Skipping file (no matching prefix): {filename}")
            skipped_files.append(filename)
    console_print("")
    console_print("=" * 70)
    console_print("📊 Batch processing completion summary:")
    console_print(f"   ✅ Successfully processed: {len(processed_files)} files")
    console_print(f"   ❌ Processing failed: {len(failed_files)} files")
    console_print(f"   ⏭️ Skipped files: {len(skipped_files)} files")
    console_print("")
    if processed_files:
        console_print("✅ Successfully processed files:")
        for file in processed_files:
            console_print(f"   • {file}")
        console_print("")
    if failed_files:
        console_print("❌ Processing failed files:")
        for file in failed_files:
            console_print(f"   • {file}")
        console_print("")
    if skipped_files:
        console_print("⏭️ Skipped files:")
        for file in skipped_files:
            console_print(f"   • {file}")
        console_print("")
    console_print(f"⏰ Processing end time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    console_print("=" * 70)

def validate_configuration():
    errors = []
    if not os.path.exists(base_directory):
        errors.append(f"Base directory does not exist: {base_directory}")
    if not file_configs:
        errors.append("No file configurations specified")
    if advanced_settings["max_retries"] < 1:
        errors.append("max_retries must be at least 1")
    if advanced_settings["retry_delay_base"] < 1:
        errors.append("retry_delay_base must be at least 1")
    return errors

def main():
    global logger
    log_filepath = None
    try:
        config_errors = validate_configuration()
        if config_errors:
            print("❌ Configuration errors found:")
            for error in config_errors:
                print(f"   • {error}")
            return 1
        logger = setup_logging()
        log_filepath = os.environ.get("log_filepath")
        process_excel_files_in_directory(base_directory, file_configs)
        console_print("")
        console_print("🎉 Program execution completed")
        console_print("=" * 80)
        return 0
    except KeyboardInterrupt:
        console_print("\n⏹️ Program interrupted by user", level='warning')
        return 1
    except Exception as e:
        console_print(f"\n💥 Unexpected error occurred during program execution: {str(e)}", level='error')
        return 1
    finally:
        if logger and log_filepath and os.path.exists(log_filepath):
            console_print("Preparing to send notification email...")
            content = ""
            try:
                with open(log_filepath, 'r', encoding='utf-8') as f:
                    content = f.read()
            except Exception as e:
                content = f"Could not read log file: {e}"
            email_body = f"""Hello Team,
This is an automated notification.

{content}

Best regards,
Your Automation Script
"""
            send_outlook_email(
                to_recipients=to_recipients,
                subject=f"{email_subject} ({time.strftime('%Y-%m-%d %H:%M:%S')})",
                body=email_body,
                cc_recipients=cc_recipients,
                bcc_recipients=bcc_recipients
            )
            console_print("📄 Notification email sent.")
        elif not log_filepath:
            print("Log file path not found, cannot send email.")
        if logger:
            console_print("📄 Log file saved successfully")

if __name__ == "__main__":
    main()
