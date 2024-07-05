# Gmail 整理工具詳細憑證設置指南

本指南將引導您完成 Gmail API 憑證的獲取和設置過程,包括常見問題的解決方法。

## 步驟 1: 創建專案和啟用 API

1. 訪問 [Google Cloud Console](https://console.cloud.google.com/)
2. 在頁面頂部創建新專案或選擇現有專案
3. 在左側菜單中,選擇 "API 和服務" > "資料庫"
4. 在搜索欄中輸入 "Gmail API" 並選擇它
5. 點擊 "啟用" 按鈕啟用 Gmail API

## 步驟 2: 設置 OAuth 同意畫面

1. 在左側菜單中,選擇 "API 和服務" > "OAuth 同意畫面"
2. 選擇 "外部" 用戶類型 (除非您有 G Suite 帳戶)
3. 填寫必要信息:
   - 應用名稱: 例如 "Gmail Organizer"
   - 用戶支持電子郵件: 您的電子郵件地址
   - 開發者聯繫信息: 您的電子郵件地址
4. 在 "範圍" 部分,點擊 "添加或刪除範圍"
5. 在彈出的窗口中,找到並選擇 `https://www.googleapis.com/auth/gmail.modify`
   - 如果找不到,可以在底部的 "手動添加範圍" 文本框中輸入此 URL
6. 保存並繼續

## 步驟 3: 創建憑證

1. 在左側菜單中,選擇 "API 和服務" > "憑證"
2. 點擊頁面頂部的 "創建憑證" > "OAuth 客戶端 ID"
3. 在 "應用類型" 下拉菜單中選擇 "桌面應用程式"
4. 為憑證命名,例如 "Gmail Organizer Windows"
5. 點擊 "創建" 按鈕

## 步驟 4: 下載和設置憑證文件

1. 創建後,會看到一個包含您的客戶端 ID 和客戶端密鑰的彈窗
2. 點擊 "下載 JSON" 按鈕
3. 文件會以 `client_secret_[一長串數字和字母].apps.googleusercontent.com.json` 的格式下載
4. 將下載的文件重命名為 `credentials.json`
5. 將 `credentials.json` 移動到您的 Gmail 整理工具 Python 腳本所在的目錄

## 注意事項

- **測試狀態**: 您的應用initially會處於 "測試" 狀態,這對個人使用來說已經足夠
- **範圍選擇**: 確保選擇了 `https://www.googleapis.com/auth/gmail.modify` 範圍,這允許讀取和修改郵件
- **應用類型**: 對於在本地 Windows 電腦上運行的腳本,選擇 "桌面應用程式" 選項
- **首次運行**: 腳本第一次運行時會打開瀏覽器要求授權,這是正常的
- **憑證保護**: 不要分享 `credentials.json` 或將其上傳到公共代碼庫

## 常見問題解決

1. **找不到 "啟用 Gmail API" 選項**:
   - 確保您在 "API 和服務" > "資料庫" 中搜索 "Gmail API"
   - 如果仍然找不到,嘗試刷新頁面或重新登錄

2. **無法找到正確的範圍**:
   - 在 OAuth 同意畫面設置中,使用 "手動添加範圍" 選項
   - 直接輸入 `https://www.googleapis.com/auth/gmail.modify`

3. **不確定選擇哪種應用類型**:
   - 對於在本地電腦上運行的 Python 腳本,始終選擇 "桌面應用程式"

如果遇到其他問題,請查看錯誤信息或聯繫支持人員尋求幫助。
