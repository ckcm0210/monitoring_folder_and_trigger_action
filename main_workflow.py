# main_workflow.py

import sys
import os

current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

try:
    import monitoring
except ImportError as e:
    print(f"Error: Could not import monitoring.py. Please ensure it's in the same directory or accessible via sys.path. Details: {e}")
    sys.exit(1)

def main_workflow_entry():
    print("--- Starting the overall file monitoring and update automation workflow ---")
    print("This workflow initiates continuous file monitoring. ")
    print("If specific file changes are detected by 'monitoring.py', ")
    print("it will automatically trigger the 'updating.py' script. ")
    print("Upon completion of the update, a notification email will be sent via 'send_outlook_email.py'.")
    print("Press Ctrl+C to stop the monitoring process at any time.\n")

    try:
        print("[INFO] Initializing file monitoring system...")
        # 調用 monitoring 模組中的 monitor_files 函數
        # 這個函數包含了所有文件監控的邏輯
        monitoring.monitor_files()
        print("\n[INFO] File monitoring process has concluded successfully (possibly due to an update trigger or internal condition).")

    except KeyboardInterrupt:
        print("\n[WARNING] Workflow stopped manually by user (KeyboardInterrupt).")
    except Exception as e:
        print(f"\n[ERROR] An unexpected error occurred during the workflow: {e}")

    print("\n--- Overall automation workflow finished. ---")

if __name__ == "__main__":
    main_workflow_entry()