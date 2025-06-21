# File Monitoring and Automated Update System  
[中文版說明 → README.md](README.md)

This project is an automation solution tailored for Windows/Excel business scenarios. It continuously monitors specified folders, intelligently triggers update scripts based on file changes, and sends email notifications with logs upon completion. It's ideal for office reports automation and cross-department document synchronization.

---

## Table of Contents
- [Project Overview](#project-overview)
- [Features](#features)
- [File Structure](#file-structure)
- [Quick Start](#quick-start)
- [Requirements & Installation](#requirements--installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [FAQ](#faq)
- [Contributing](#contributing)
- [License](#license)

---

## Project Overview

This project enables:
- Continuous monitoring of multiple folders for file changes, based on customizable rules (Group A/Group B).
- When files in Group A are newer than those in Group B and pass a cooldown check, it automatically triggers an update script (e.g., refreshing Excel, running macros, updating links).
- After processing, the system automatically sends a detailed log email (via Outlook) to configured recipients.
- All key parameters can be flexibly adjusted in YAML configuration files.

---

## Features

- **Multi-folder monitoring**: Monitor multiple directories simultaneously.
- **Intelligent trigger rules**: Group A/B rules support keywords or regex for matching.
- **Automatic cooldown**: Ensures files are stable before processing.
- **Excel automation**: Handles password-protected files, runs macros, refreshes links/connections.
- **Detailed logs & email notification**: Automatically sends logs to preset recipients.
- **Highly configurable**: All rules, paths, email, passwords, and advanced options are set in YAML files.
- **Error handling & notification**: Robust exception handling and email alerts on failure.

---

## File Structure

```
your_project_root/
├── main_workflow.py               # Main workflow entry point
├── monitoring.py                  # File monitoring logic
├── monitoring_config.yaml         # Monitoring parameters
├── updating.py                    # Excel automation update script
├── updating_config.yaml           # Update script parameters
├── utility/
│   └── send_outlook_email.py      # Outlook email utility
├── !_run_me_to_start_monitoring.ipynb # Jupyter start example
├── LICENSE                        # License (MIT)
├── README.md                      # Chinese documentation
└── README_en.md                   # This file
```

---

## Quick Start

1. **Clone the repository**
    ```bash
    git clone https://github.com/ckcm0210/monitoring_folder_and_trigger_action.git
    cd monitoring_folder_and_trigger_action
    ```

2. **Install dependencies**
    If you have a requirements.txt:
    ```bash
    pip install -r requirements.txt
    ```
    If not, install manually:
    ```bash
    pip install pywin32 pyyaml openpyxl
    ```

3. **Edit configuration files**  
    - Copy and edit `monitoring_config.yaml` and `updating_config.yaml` for your environment (paths, file patterns, email recipients, etc).
    - It's recommended to keep example config files for future reference.

4. **Check your environment**
    - Windows only (requires Excel and Outlook automation).
    - Make sure Outlook and Excel are installed and can be launched.

---

## Requirements & Installation

- **Operating System**: Windows only
- **Python**: 3.7+
- **Required Packages**:
    - pywin32 (for Outlook/Excel automation)
    - openpyxl (Excel file handling)
    - PyYAML (YAML config parsing)
- **Microsoft Outlook & Excel must be installed**

---

## Configuration

### 1. `monitoring_config.yaml`
- Defines which folders to monitor, file matching rules, and which update scripts to run.
- Group A/B supports regex or keyword matching, cooldown/intervals are adjustable.

#### Example
```yaml
folders:
  - folder_path: "K:\\Chain\\2024Q4\\Preliminary\\Test2 - new"
    updating_script: "V:\\SomeFolder\\updating.py"
file_group_a:
  - "Data - Section"
  - "Data - Taxes"
file_group_b:
  - "Data - All"
  - "Chain Summary"
check_interval: 2
cooldown_period: 2
```

### 2. `updating_config.yaml`
- Sets email recipients, Excel passwords/macros, log directory, and advanced automation options.

#### Example
```yaml
email_recipients:
  to: ["your_email@example.com"]
  cc: []
  bcc: []
file_configs:
  Data - All:
    macro: null
    open_password: null
    write_password: "abc123"
advanced_settings:
  max_retries: 3
  retry_delay_base: 2
  excel_visible: False
  force_calculation: True
```

---

## Usage

1. **Start the monitoring workflow:**
    ```bash
    python main_workflow.py
    ```

2. **Jupyter Notebook (optional):**
    Open and run `!_run_me_to_start_monitoring.ipynb`.

3. **To stop monitoring:**
    - Press `Ctrl+C` in the terminal/command prompt.

---

## FAQ

### 1. Why can't it send email automatically?
- Make sure Outlook is installed, running, and your account is logged in.
- Check that Python has permission to control Outlook.

### 2. ImportError on execution?
- Check that all filenames/paths are correct (no extra spaces).
- Ensure all required Python packages are installed.

### 3. Monitoring doesn't trigger automation?
- Check if your Group A/B rules correctly match your files.
- Too short a cooldown may cause files to be processed before they are finished.

### 4. Excel password/macro issues?
- Make sure you have the correct password/macro name set in the config. Adjust macro security/trusted locations as needed.

For more issues, please open an issue on GitHub!

---

## Contributing

- Issues and pull requests for bug fixes or feature suggestions are welcome.
- Please do not commit sensitive information (like passwords) directly to the public repository.
- Read the code comments and documentation before contributing.

---

## License

This project is licensed under the [MIT License](LICENSE). You are free to use, modify, and distribute it.

---
