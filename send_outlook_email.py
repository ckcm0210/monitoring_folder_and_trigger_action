import win32com.client # 用於與 Outlook 應用程式溝通
import sys            # 用於程式出錯時的安全退出
import os             # 用於檢查檔案是否存在
import time           # 用於模擬任務延遲

##
# `send_outlook_email` 函數定義
#
def send_outlook_email(to_recipients: list, subject: str, body: str = "", html_body: str = None, 
                       attachments: list = None, cc_recipients: list = None, bcc_recipients: list = None) -> bool:
    """
    透過 Outlook 發送電郵，支援 TO、CC、BCC 的列表輸入，並可包含純文字、HTML 內容及附件。

    參數:
        to_recipients (list): 主要收件人嘅電郵地址列表。
        subject (str): 電郵主旨。
        body (str): 電郵嘅純文字內容。預設為空字串。
        html_body (str, optional): 電郵嘅 HTML 內容。如果提供，將覆蓋 'body' 內容。預設為 None。
        attachments (list, optional): 附件檔案路徑嘅列表。預設為 None。
        cc_recipients (list, optional): CC 收件人嘅電郵地址列表。預設為 None。
        bcc_recipients (list, optional): BCC 收件人嘅電郵地址列表。預設為 None。

    返回 (bool):
        如果電郵成功發送，返回 True；否則返回 False。
    """
    app_outlook = None # 初始化 app_outlook 變數，以確保 finally 區塊能安全檢查

    try:
        # 嘗試連接到 Outlook 應用程式
        app_outlook = win32com.client.Dispatch("Outlook.Application")
        print("Successfully connected to Outlook application.")
    except Exception as e:
        print(f"Error: Could not connect to Outlook application. Please ensure Outlook is installed and running.")
        print(f"Detailed error message: {e}")
        return False # 連接失敗，直接返回 False

    print("Preparing to send email...")
    try:
        mail_item = app_outlook.CreateItem(0) # 0 代表 olMailItem，即一個郵件物件

        # 處理 TO 收件人列表
        if to_recipients:
            mail_item.To = "; ".join(to_recipients)
            print(f"To recipients set to: {mail_item.To}")
        else:
            print("Warning: No 'To' recipients provided. Email might not be sent or might require manual input.")

        # 處理 CC 收件人列表
        if cc_recipients:
            mail_item.CC = "; ".join(cc_recipients)
            print(f"CC recipients set to: {mail_item.CC}")

        # 處理 BCC 收件人列表
        if bcc_recipients:
            mail_item.BCC = "; ".join(bcc_recipients)
            print(f"BCC recipients set to: {mail_item.BCC}")

        # 設定郵件主旨
        mail_item.Subject = subject
        print(f"Subject set to: '{subject}'")

        # 設定郵件內容 (優先使用 HTML 內容)
        if html_body:
            mail_item.HTMLBody = html_body
            print("Email HTML body content set.")
        else:
            mail_item.Body = body
            print("Email plain text body content set.")

        # 添加附件
        if attachments:
            for attachment_path in attachments:
                if os.path.exists(attachment_path):
                    mail_item.Attachments.Add(attachment_path)
                    print(f"Attachment added: {attachment_path}")
                else:
                    print(f"Warning: Attachment file '{attachment_path}' not found, skipping.")

        # 發送郵件
        mail_item.Send()
        print("Email sent successfully! Please check your inbox and sent items folder.")
        return True # 發送成功，返回 True

    except Exception as e:
        print(f"Error: Failed to send email. Check Outlook settings or permissions.")
        print(f"Detailed error message: {e}")
        return False # 發送失敗，返回 False
    finally:
        # 釋放 COM 物件，幫助系統回收資源
        if 'mail_item' in locals() and mail_item is not None:
            del mail_item
            print("mail_item object released.")
        if 'app_outlook' in locals() and app_outlook is not None:
            del app_outlook
            print("app_outlook object released.")
        print("\nProgram finished. All COM objects attempted to be released.")

---

##
# `if __name__ == "__main__":` 主程式模板
#
if __name__ == "__main__":
    print("--- Script Started ---")

    # --- Step 1: Define your main task(s) here ---
    # 你可以在這裡寫你的 Python 程式碼，執行你主要的自動化任務。
    # 例如：讀取 Excel 檔案、處理數據、生成報告、從資料庫獲取資料等等。
    print("Executing your main task(s)...")
    
    # --- 模擬你的主要任務輸出 ---
    # 假設你的任務會生成一些需要放到電郵裡的資訊和可能需要附加的檔案。
    
    # 模擬生成電郵內文
    # task_completion_summary 會儲存任務完成後需要通知嘅文字內容
    task_completion_summary = "All daily data processing and report generation tasks have been completed successfully. The system is stable."

    # 模擬一個需要附加的檔案路徑
    # report_attachment_path 會儲存要附加嘅檔案路徑
    report_attachment_path = r'C:\Users\YourUser\Documents\Daily_Report_2025_06_12.pdf' # <-- **請替換為你的實際檔案路徑！**
                                                                                   # 例如：r'C:\MyReports\SalesData.xlsx'

    # 為了演示，我們創建一個虛擬文件（在實際應用中，這個文件應該是你的任務生成的）
    if not os.path.exists(os.path.dirname(report_attachment_path)):
        os.makedirs(os.path.dirname(report_attachment_path)) # 確保目錄存在
    try:
        with open(report_attachment_path, 'w', encoding='utf-8') as f:
            f.write(f"This is a simulated daily report content generated on {time.strftime('%Y-%m-%d %H:%M:%S')}.\n")
            f.write(f"Task summary: {task_completion_summary}\n")
            f.write("More detailed information can be found in the actual report system.\n")
        print(f"Simulated attachment created at: {report_attachment_path}")
    except Exception as e:
        print(f"Error creating simulated attachment: {e}")
        report_attachment_path = None # 如果創建失敗，則不添加此附件

    print("Main task(s) completed.")

    # --- Step 2: Prepare email details ---
    # 這裡準備要發送電郵的所有資訊，包括收件人、主旨、內容和附件。
    
    # 主要收件人：你可以放多個電郵地址在列表中
    # to_recipients_list 呢個變數會儲存主要收件人嘅電郵地址列表
    to_recipients_list = ["your_email@example.com", "your.colleague@example.com"] # <-- **請替換為你的實際收件人電郵地址！**

    # CC 收件人：可選，如果沒有 CC 人員，可以留空列表 []
    # cc_recipients_list 呢個變數會儲存 CC 收件人嘅電郵地址列表
    cc_recipients_list = ["your.manager@example.com"] # <-- **可選：替換為你的實際 CC 人員電郵地址，或設為 []**

    # BCC 收件人：可選，如果沒有 BCC 人員，可以留空列表 []
    # bcc_recipients_list 呢個變數會儲存 BCC 收件人嘅電郵地址列表
    bcc_recipients_list = [] # <-- **可選：替換為你的實際 BCC 人員電郵地址，或設為 []**

    # 電郵主旨
    # email_subject 呢個變數會儲存電郵嘅主旨
    email_subject = f"Automated Daily Task Completion Notification ({time.strftime('%Y-%m-%d')})"

    # 電郵內容：可以是你任務生成的摘要，也可以是固定文字
    # email_body 呢個變數會儲存電郵嘅純文字內容
    email_body = f"""Hello Team,

This is an automated notification.

{task_completion_summary}

Please find the detailed report attached.

Best regards,
Your Automation Script
"""
    # 如果你想發送 HTML 格式的電郵，可以像下面這樣定義 html_body
    # email_html_body = f"""
    # <html>
    # <body style="font-family: Arial, sans-serif;">
    #     <h2 style="color: #4CAF50;">Automated Daily Task Update</h2>
    #     <p>Hello Team,</p>
    #     <p>This is an automated notification.</p>
    #     <p style="background-color: #e0ffe0; padding: 10px; border-radius: 5px;">
    #         <b>Summary:</b> {task_completion_summary}
    #     </p>
    #     <p>Please find the detailed report attached for your review.</p>
    #     <p>Best regards,<br>Your Automation Script</p>
    # </body>
    # </html>
    # """

    # 附件列表：如果你的任務生成了檔案需要附加，將其路徑加入此列表
    # attachments_to_send 呢個變數會儲存所有要附加嘅檔案路徑列表
    attachments_to_send = []
    if report_attachment_path and os.path.exists(report_attachment_path):
        attachments_to_send.append(report_attachment_path)
    else:
        print(f"Warning: Report attachment '{report_attachment_path}' not found. Email will be sent without this attachment.")

    # --- Step 3: Send the email ---
    print("\nAttempting to send notification email...")
    
    # 調用 send_outlook_email 函數，傳入所有準備好的參數
    email_sent_successfully = send_outlook_email(
        to_recipients=to_recipients_list,
        subject=email_subject,
        body=email_body,
        # html_body=email_html_body, # 如果要發送 HTML 郵件，請取消註釋此行並註釋掉上面的 body 參數
        attachments=attachments_to_send,
        cc_recipients=cc_recipients_list,
        bcc_recipients=bcc_recipients_list
    )

    # --- Step 4: Handle email sending result ---
    if email_sent_successfully:
        print("Email notification successfully processed.")
    else:
        print("Email notification failed to send. Please check the console output above for detailed errors.")

    print("--- Script Finished ---")