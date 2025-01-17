import os
import pickle
import json
import webbrowser
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk
import threading

class GmailAuthWizard:
    def __init__(self):
        self.SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
        self.creds_file = 'credentials.json'
        self.token_file = 'token.pickle'
        
        self.setup_ui()
        
    def setup_ui(self):
        """設置使用者介面"""
        self.window = tk.Tk()
        self.window.title("Gmail 授權精靈")
        self.window.geometry("600x450")
        
        # 使用 Frame 來組織元件
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 說明文字
        ttk.Label(main_frame, text="Gmail 授權精靈", font=('Arial', 14, 'bold')).grid(row=0, column=0, columnspan=2, pady=10)
        
        # 重要提醒
        warning_frame = ttk.LabelFrame(main_frame, text="重要提醒", padding="10")
        warning_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        warning_text = """在開始授權之前，請確保您已完成以下步驟：

1. 打開 Google Cloud Console
2. 進入您的專案
3. 選擇 'OAuth 同意畫面'
4. 在 '測試使用者' 區段中，點擊 '+ 新增使用者'
5. 輸入您要授權的 Gmail 帳號
6. 儲存設定

只有被加入為測試使用者的 Gmail 帳號才能使用此應用程式。"""
        
        warning_label = ttk.Label(warning_frame, text=warning_text, wraplength=500, justify="left")
        warning_label.grid(row=0, column=0, pady=5)
        
        # 新增開啟 Google Cloud Console 按鈕
        ttk.Button(warning_frame, 
                  text="開啟 Google Cloud Console", 
                  command=lambda: webbrowser.open('https://console.cloud.google.com')
        ).grid(row=1, column=0, pady=5)
        
        # Gmail 帳號輸入區
        input_frame = ttk.LabelFrame(main_frame, text="授權設定", padding="10")
        input_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        ttk.Label(input_frame, text="請輸入要授權的 Gmail 帳號：").grid(row=0, column=0, pady=5)
        
        self.email_var = tk.StringVar()
        email_entry = ttk.Entry(input_frame, textvariable=self.email_var, width=40)
        email_entry.grid(row=1, column=0, pady=5)
        
        # 狀態標籤
        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.grid(row=3, column=0, columnspan=2, pady=10)
        
        # 進度條
        self.progress = ttk.Progressbar(main_frame, orient="horizontal", length=300, mode="indeterminate")
        self.progress.grid(row=4, column=0, columnspan=2, pady=10)
        
        # 開始按鈕
        self.start_button = ttk.Button(main_frame, text="開始授權", command=self.start_auth)
        self.start_button.grid(row=5, column=0, columnspan=2, pady=10)

    def update_status(self, message, is_error=False):
        """更新狀態訊息"""
        self.status_label.config(text=message, foreground="red" if is_error else "black")
        self.window.update()

    def handle_auth_error(self, error):
        """處理授權錯誤"""
        if "access_denied" in str(error).lower():
            messagebox.showerror("授權錯誤", 
                               "存取被拒絕。可能的原因：\n\n"
                               "1. 此 Gmail 帳號尚未被加入測試使用者清單\n"
                               "2. 您在授權過程中點擊了 '取消' 或 '拒絕'\n\n"
                               "請確保您的 Gmail 帳號已被加入測試使用者清單，然後再試一次。")
        else:
            messagebox.showerror("錯誤", f"授權過程中發生錯誤：\n{str(error)}")

    def start_auth(self):
        """開始授權流程"""
        email = self.email_var.get().strip()
        if not email or '@' not in email:
            messagebox.showerror("錯誤", "請輸入有效的 Gmail 帳號")
            return
        
        self.start_button.config(state='disabled')
        self.progress.start()
        
        # 在新執行緒中執行授權流程
        threading.Thread(target=self.auth_process, args=(email,), daemon=True).start()

    def auth_process(self, email):
        """執行授權流程"""
        try:
            self.update_status("檢查憑證檔案...")
            
            if not os.path.exists(self.creds_file):
                self.update_status("錯誤：找不到 credentials.json 檔案", True)
                messagebox.showerror("錯誤", 
                                   "找不到 credentials.json 檔案！\n\n"
                                   "請確保：\n"
                                   "1. 您已從 Google Cloud Console 下載憑證檔案\n"
                                   "2. 檔案已重命名為 credentials.json\n"
                                   "3. 檔案位於程式同一目錄下")
                self.reset_ui()
                return
            
            creds = None
            if os.path.exists(self.token_file):
                self.update_status("載入既有的授權資訊...")
                with open(self.token_file, 'rb') as token:
                    creds = pickle.load(token)
            
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    self.update_status("更新過期的授權資訊...")
                    creds.refresh(Request())
                else:
                    self.update_status("開始新的授權流程...")
                    flow = InstalledAppFlow.from_client_secrets_file(
                        self.creds_file, self.SCOPES)
                    creds = flow.run_local_server(port=0)
                
                self.update_status("儲存授權資訊...")
                with open(self.token_file, 'wb') as token:
                    pickle.dump(creds, token)
            
            self.update_status("授權成功！")
            messagebox.showinfo("成功", f"Gmail 帳號 {email} 已成功授權！")
            
        except Exception as e:
            self.handle_auth_error(e)
        
        finally:
            self.reset_ui()

    def reset_ui(self):
        """重置 UI 狀態"""
        self.progress.stop()
        self.start_button.config(state='normal')

    def run(self):
        """運行程式"""
        self.window.mainloop()

if __name__ == "__main__":
    wizard = GmailAuthWizard()
    wizard.run()