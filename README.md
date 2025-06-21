# 文件監控與自動化更新系統

這是一個自動化的 Python 解決方案，旨在監控特定資料夾中的文件變更。當檢測到預設文件組（Group A）比另一組文件（Group B）更新時，它將自動觸發一個更新腳本，並在更新完成後發送電子郵件通知。

## 專案概述

本專案主要由以下幾個模組組成：

* **`main_workflow.py`**: 整個自動化流程的入口點。它負責初始化文件監控系統，並在監控過程中協調各個模組的執行。
* **`monitoring.py`**: 核心文件監控模組。它會定期檢查配置中定義的資料夾，比較 Group A 和 Group B 文件組的最新修改時間，以判斷是否需要觸發更新。
* **`updating.py`**: 當 `monitoring.py` 檢測到文件變更並滿足觸發條件時，會調用此腳本執行實際的更新操作。此腳本支援對 Excel 檔案進行操作（如執行宏、處理密碼），並會生成操作日誌。
* **`send_outlook_email.py`**: 一個實用工具模組，用於透過 Outlook 發送電子郵件通知。它被 `updating.py` 用於在更新完成後發送帶有日誌內容的通知郵件。
* **`monitoring_config.yaml`**: 監控模組的配置檔案，定義了要監控的資料夾、文件組的正則表達式模式、檢查間隔和冷卻時間等。
* **`updating_config.yaml`**: 更新模組的配置檔案，定義了日誌目錄、Excel 檔案的宏和密碼設定、郵件接收者和主旨前綴等。

## 檔案結構

為了確保程式能正確地找到並導入所有模組和配置檔案，建議您的專案目錄結構如下：

your_project_root/
├── main_workflow.py
├── monitoring.py
├── monitoring_config.yaml
├── updating.py
├── updating_config.yaml
└── utility/
     └── send_outlook_email.py

* `your_project_root/`: 這是您的專案主目錄。
* `main_workflow.py`、`monitoring.py`、`updating.py`：這些是主要的 Python 腳本，直接位於專案根目錄下。
* `monitoring_config.yaml`、`updating_config.yaml`：這些是配置檔案，也直接位於專案根目錄下。
* `utility/`：這是一個子目錄，用於存放輔助性質的 Python 模組。
* `send_outlook_email.py`：這個模組位於 `utility/` 子目錄中，被其他主腳本導入使用。

這種結構符合 Python 的模組化最佳實踐，並便於管理。

## 系統功能

* **持續文件監控**: 定期掃描預設資料夾。
* **智能變更檢測**: 根據 Group A 和 Group B 文件的修改時間判斷是否觸發更新。
* **穩定性冷卻期**: 在文件變動後，等待一段冷卻時間，確保文件穩定後才執行更新，避免在文件仍在寫入時觸發。
* **自動化更新觸發**: 滿足條件時自動執行預設的更新腳本。
* **Excel 自動化**: `updating.py` 能夠處理 Excel 檔案，包括開啟受密碼保護的檔案、執行 VBA 宏、保存並關閉檔案。
* **詳細日誌記錄**: 記錄所有重要的操作和潛在問題。
* **郵件通知**: 在更新流程完成後，透過 Outlook 發送詳細的操作日誌作為郵件內容。
* **可配置性**: 大部分關鍵設定都儲存在 `.yaml` 配置文件中，方便用戶根據需求修改。

## 環境要求

1.  **Python 環境**: 確保您的系統安裝了 Python 3.x。
2.  **依賴庫**:
    * **推薦安裝方式**: 在專案根目錄下運行 `pip install -r requirements.txt`。
    * 手動安裝（如果沒有 `requirements.txt`）:
        * `pywin32` (`pip install pywin32`)：用於與 Outlook 應用程式互動。
        * `openpyxl` (`pip install openpyxl`)：用於處理 Excel XLSX/XLSM 檔案。
        * `PyYAML` (`pip install PyYAML`)：用於讀取 `config.yaml` 檔案。
3.  **Outlook 應用程式**: `send_outlook_email.py` 腳本依賴於 Outlook 應用程式的運行。

## 安裝

1.  **克隆倉庫**:
    ```bash
    git clone <您的_GitHub_倉庫地址>
    cd <您的_專案資料夾名稱>
    ```

2.  **安裝依賴**:
    推薦使用 `pip` 安裝所有必要的 Python 庫。
    ```bash
    pip install pywin32 pyyaml openpyxl
    ```
    * `pywin32`: 提供了 `win32com.client`，用於與 Windows 應用程式（如 Outlook 和 Excel）交互。
    * `pyyaml`: 用於讀取和解析 `.yaml` 配置文件。
    * `openpyxl`: 用於處理 `.xlsx` 格式的 Excel 檔案（`updating.py` 中可能用到，儘管主要依賴 `win32com`）。

## 配置

在運行程式之前，您需要配置以下兩個 `.yaml` 檔案以符合您的環境和需求：

1.  **`monitoring_config.yaml`**:
    * `folders`: 定義要監控的資料夾路徑及其對應的更新腳本路徑。
    * `file_group_a`: 一組正則表達式模式，用於匹配需要監控的文件（通常是數據輸入或更新文件）。
    * `file_group_b`: 另一組正則表達式模式，用於匹配作為比較基準的文件（通常是輸出或彙總文件）。
    * `check_interval`: 每次檢查文件變更之間的等待時間（秒）。
    * `cooldown_period`: 在文件穩定後，執行更新腳本前的冷卻時間（秒）。

    範例：
    ```yaml
    folders:
      - folder_path: "K:\\Chain\\2024Q4\\Preliminary\\Test2 - new"
        updating_script: "V:\\新增資料夾\\updating.py" # 注意：這裡的 updating_script 路徑仍指向實際的 updating.py 檔案
      - folder_path: "K:\\Chain\\2024Q4\\Preliminary\\Test2 - interim"
        updating_script: "V:\\新增資料夾\\updating.py" # 同上
    file_group_a:
      - "Data - Section"
      - "Data - Taxes"
      - "Data - Ownership"
      - "Related"
    # ... 其他配置
    ```

2.  **`updating_config.yaml`**:
    * `email_recipients`: 配置郵件的 `to`、`cc` 和 `bcc` 收件人列表。
    * `email_subject_prefix`: 郵件主旨的前綴。
    * `log_directory`: 程式運行日誌文件的儲存目錄。
    * `file_configs`: 針對特定 Excel 檔案的配置，例如宏名稱、開啟密碼和寫入密碼。
    * `advanced_settings`: 高級設定，例如 Excel 應用程式是否可見、重試次數等。

    範例：
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
    # ... 其他配置
    ```
    **注意**: `updating_script` 在 `monitoring_config.yaml` 中指定的路徑，應指向實際的 `updating.py` 檔案。

## 如何運行

確保您已經完成了上述的安裝和配置步驟。

1.  **確保 Outlook 運行中**: 由於郵件功能依賴於 Outlook，請確保您的 Outlook 應用程式正在運行。
2.  **運行主工作流程**:
    打開您的命令提示符或終端機，導航到專案的根目錄，然後運行 `main_workflow.py`：
    ```bash
    python main_workflow.py
    ```

    或者，如果您是在 Jupyter Notebook 環境中，可以直接運行 `main_workflow_entry()` 函數：
    ```python
    # 在 Jupyter Notebook 中
    import sys
    import os

    current_dir = os.path.dirname(os.path.abspath(__file__))
    # 將專案根目錄加入 sys.path，確保可以找到 utility 模組
    project_root = os.path.abspath(os.path.join(current_dir)) # 確保 project_root 是當前檔案所在的目錄
    if project_root not in sys.path:
        sys.path.append(project_root)

    import main_workflow
    main_workflow.main_workflow_entry()
    ```

程式將會開始監控文件。您可以隨時按下 `Ctrl+C` 來手動停止監控流程。

## 注意事項

* **檔案路徑**: 請確保所有配置中的檔案和資料夾路徑都是正確且可訪問的。
* **Outlook 和 Excel 權限**: 確保運行腳本的用戶帳戶有足夠的權限來控制 Outlook 和 Excel 應用程式。
* **日誌目錄**: 確保 `updating_config.yaml` 中配置的 `log_directory` 存在且可寫。
* **Excel 密碼**: 如果 Excel 檔案有密碼保護，請在 `updating_config.yaml` 中正確配置 `open_password` 或 `write_password`。
* **宏安全性**: 如果您的 Excel 檔案包含宏，請確保 Excel 的宏安全設定允許執行宏，或者將您的專案資料夾添加到信任位置。
* **文件名空格**: 請確保所有導入的 Python 檔案名不包含空格，例如 `send_outlook_email.py` 而不是 `send_outlook_email .py`。

## 故障排除

* **`ImportError`**: 如果運行時遇到 `ImportError`，請檢查檔案名稱是否正確，並確保所有必要的模組都已安裝。特別是檢查 `send_outlook_email.py` 的檔案名稱是否已修正（已從 `send_outlook_email .py` 修改為 `send_outlook_email.py`），且其路徑符合 `utility/send_outlook_email.py`。
* **Outlook/Excel 自動化問題**:
    * 確保 Outlook 和 Excel 已安裝並正常運行。
    * 檢查是否有任何安全彈出窗口阻止了程式的交互。
    * 檢查 `pywin32` 是否正確安裝。
* **文件未觸發更新**:
    * 檢查 `monitoring_config.yaml` 中的 `file_group_a` 和 `file_group_b` 正則表達式是否正確匹配您的文件。
    * 檢查文件的修改時間是否符合預期，即 Group A 確實比 Group B 新。
    * 調整 `check_interval` 和 `cooldown_period` 以適應您的需求。

## 貢獻

如果您有任何建議或發現錯誤，歡迎提出 issue 或 pull request。

## 許可證

[在此處放置您的許可證信息，例如 MIT 許可證]
