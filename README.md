# Gmail 郵件整理工具

自動下載 Gmail 附件並刪除舊郵件的 Python 腳本。

## 快速開始

1. 安裝 Python 3.6+
2. 安裝依賴：
   ```
   pip install google-auth-oauthlib google-auth-httplib2 google-api-python-client
   ```
3. 設置 Google Cloud Project：
   - 訪問 [Google Cloud Console](https://console.cloud.google.com/)
   - 創建項目並啟用 Gmail API
   - 下載 OAuth 2.0 客戶端 ID，重命名為 `credentials.json`

## 使用方法

1. 運行腳本創建配置文件：
   ```
   python gmail_organizer.py
   ```
2. 編輯 `config.json` 設置標籤、日期範圍等
3. 再次運行腳本處理郵件：
   ```
   python gmail_organizer.py
   ```

首次運行時，瀏覽器會要求授權訪問您的 Gmail。

## 注意事項

- 保護好 `credentials.json` 和 `token.json`
- 定期檢查 Google 帳戶活動

需要幫助？提出 issue 或 pull request。

