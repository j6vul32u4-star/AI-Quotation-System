---
alwaysApply: true
---

# 報價系統安全與開發規範 (Quote Safety Rules)

1. **嚴禁揭露成本底價**：GAS (Google Apps Script) 讀取庫存表時，必須在資料處理階段剔除 `cost_price` 成本欄位，嚴禁在任何回傳結果、日誌或訊息中揭露底價。
2. **業務白名單驗證**：所有透過 LINE 進行的查詢請求，必須先驗證發訊者是否位於 `ALLOWED_SALES_IDS` 業務白名單內；未授權帳號一律拒絕存取。
3. **台灣機械業繁體中文術語**：系統回應、註解、UI 文字與報價單內容，一律使用台灣機械業慣用之繁體中文（zh-TW）專業術語。
4. **自動同步雲端**：每次修改本地 `Code.js` 後，必須自動執行 `clasp push -f` 將最新程式碼強制同步至 Google Apps Script 雲端專案。
