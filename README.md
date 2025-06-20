# Automated File Update and Monitoring System

## Project Overview

This project is designed to automatically monitor changes in files within specified folders and, when certain conditions are met, trigger a Python script (`updating.py`) to perform file update operations. The system incorporates a dual-layer error notification mechanism to ensure timely alerts in case of any anomalies.

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

# 檔案自動更新與監控系統

## 專案概覽

本專案旨在自動監控指定資料夾中的檔案變動，並在符合特定條件時觸發一個 Python 腳本（`updating.py`）來執行檔案更新操作。系統設計了雙重錯誤通知機制，以確保在任何情況下都能及時收到異常警報。

### 專案文件

* **`monitoring_folder_and_trigger_action.py`**:
    主監控腳本，負責監聽資料夾中的檔案變化，判斷是否滿足觸發條件，並執行 `updating.py`。
* **`updating.py`**:
    執行實際檔案更新邏輯的腳本，包括 Excel 檔案處理、巨集執行和郵件通知。
* **`send_outlook_email.py`**:
    一個通用的輔助腳本，用於發送 Outlook 電子郵件。
* **`config.yaml`**:
    外部化設定檔，包含所有可配置的路徑、檔案規則、郵件列表等。

## 設定 (Configuration)

所有專案相關的設定都集中在 `config.yaml` 檔案中。請務必在啟動專案前仔細審查並修改此檔案以符合您的環境。

### `config.yaml` 範例及說明

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
