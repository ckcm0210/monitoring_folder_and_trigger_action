# 文件監控與自動自動化更新系統  
[English version available → README_en.md](README_en.md)

本專案是一個針對 Windows/Excel 報表流程設計的自動化解決方案。能自動監控指定資料夾，根據文件變動智能觸發更新腳本，並於流程完成後自動發送郵件通知，適合辦公室自動報表、跨部門文件同步等自動化需求。

---

## 目錄
- [專案簡介](#專案簡介)
- [特色功能](#特色功能)
- [檔案結構說明](#檔案結構說明)
- [快速開始](#快速開始)
- [環境與安裝](#環境與安裝)
- [設定說明](#設定說明)
- [使用方式](#使用方式)
- [常見問題 FAQ](#常見問題-faq)
- [貢獻方式](#貢獻方式)
- [授權條款](#授權條款)

---

## 專案簡介

此專案能夠：
- 持續監控多個資料夾內的文件，根據自訂規則（Group A/Group B）判斷是否需要自動觸發更新。
- 當檢測到 Group A 文件比 Group B 新，且通過冷卻期，會自動執行更新腳本（如自動刷新 Excel、執行巨集、重新整理連結）。
- 程序完成後，自動寄發詳細日誌郵件（經由 Outlook）給指定收件人。
- 所有關鍵參數皆可於 YAML 設定檔彈性調整。

---

## 特色功能

- **多資料夾監控**：可同時監控多個指定目錄。
- **智能觸發條件**：Group A/B 文件規則可用關鍵字或正則表達式自訂。
- **自動冷卻期**：避免檔案尚未穩定即觸發後續處理。
- **Excel 自動化**：支援密碼保護、巨集執行、連結/連線刷新。
- **詳細日誌與自動郵件**：流程完成自動寄送日誌至指定信箱。
- **高度可設定**：所有規則、路徑、郵件、密碼、進階行為皆可於 YAML 配置。
- **錯誤處理與通知**：多層次例外處理，失敗時自動郵件通知。

---

## 檔案結構說明

```
your_project_root/
├── main_workflow.py               # 主流程入口
├── monitoring.py                  # 檔案監控邏輯
├── monitoring_config.yaml         # 監控參數設定
├── updating.py                    # Excel 自動化更新腳本
├── updating_config.yaml           # 更新腳本參數設定
├── utility/
│   └── send_outlook_email.py      # Outlook 郵件寄送工具
├── !_run_me_to_start_monitoring.ipynb # Jupyter 啟動範例
├── LICENSE                        # 授權條款 (MIT)
├── README.md                      # 本說明文件
└── README_en.md                   # 英文說明文件
```

---

## 快速開始

1. **Clone 專案**
    ```bash
    git clone https://github.com/ckcm0210/monitoring_folder_and_trigger_action.git
    cd monitoring_folder_and_trigger_action
    ```

2. **安裝必要套件**
    若有 requirements.txt：
    ```bash
    pip install -r requirements.txt
    ```
    若無，請手動安裝：
    ```bash
    pip install pywin32 pyyaml openpyxl
    ```

3. **編輯設定檔**  
    - 複製並修改 `monitoring_config.yaml` 及 `updating_config.yaml`，填入你的路徑、檔名規則、信箱等資訊。
    - 建議可保留 example 配置檔以利日後參考。

4. **確認環境**
    - 僅支援 Windows（因需自動化 Excel/Outlook）。
    - 請確認 Outlook 與 Excel 已安裝並可正常啟動。

---

## 環境與安裝

- **作業系統**：僅支援 Windows
- **Python 版本**：3.7+
- **必要套件**：
    - pywin32（與 Outlook/Excel 溝通）
    - openpyxl（Excel 檔案處理）
    - PyYAML（讀取 YAML 設定）
- **需安裝 Microsoft Outlook & Excel**

---

## 設定說明

### 1. `monitoring_config.yaml`
- 設定監控哪些資料夾、匹配哪些檔案、觸發哪些更新腳本。
- Group A/Group B 支援正則或關鍵字，冷卻期可調整。

#### 範例片段
```yaml
folders:
  - folder_path: "K:\\2024Q4\\Test2 - new"
    updating_script: "V:\\新增資料夾\\updating.py"
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
- 設定郵件收件人、Excel 密碼/巨集、日誌目錄、進階自動化細節。

#### 範例片段
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

## 使用方式

1. **啟動監控主程式：**
    ```bash
    python main_workflow.py
    ```

2. **Jupyter Notebook 啟動（可選）：**
    直接開啟`!_run_me_to_start_monitoring.ipynb`並執行即可。

3. **停止監控：**
    - 於命令列視窗按 `Ctrl+C`。

---

## 常見問題 FAQ

### 1. 為什麼郵件無法自動寄出？
- 請確認 Outlook 已安裝並啟動，且帳號已登入。
- 檢查 Python 是否有權限操作 Outlook。

### 2. 執行時出現 ImportError？
- 檢查路徑與檔名（不得有多餘空白）。
- 確認所有必要 Python 套件已安裝。

### 3. 監控流程未自動觸發？
- 檢查 Group A/B 規則是否正確匹配檔案。
- 冷卻期設太短可能導致檔案未寫完即被處理。

### 4. Excel 檔案有密碼或巨集問題？
- 請於設定檔內正確填寫密碼/巨集名稱，必要時調整信任位置與巨集安全性。

更多問題歡迎開 issue 提問！

---

## 貢獻方式

- 歡迎 issue、pull request 修正錯誤或建議新功能。
- 建議勿將含敏感資訊（如帳號密碼）直接 commit 至公開 repository。
- 貢獻前請詳閱程式註解與說明文件。

---

## 授權條款

本專案採用 [MIT License](LICENSE) 授權，歡迎自由使用、修改與散布。

---
