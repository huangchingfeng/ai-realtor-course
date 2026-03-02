# n8n 工作流模板 — AI 超級房仲實戰班

> 講師：阿峰老師（黃敬峰）｜AutoLab AI 實戰學院

## 工作流總覽

| 編號 | 名稱 | 觸發方式 | 用途 | nodes 數 |
|------|------|----------|------|----------|
| W1 | 物件行銷自動產出 | Webhook（接收物件資料） | 一次產出 591 標題/描述、LINE 圖卡、FB 貼文、IG 輪播、Google 商家等 6 路行銷素材 | 15 |
| W2 | 帶看跟進提醒 | 排程（每天 08:00） | 自動檢查 Day3/Day7/Day14 到期客戶，AI 產出跟進訊息並提醒 | 11 |
| W3 | 月報自動生成 | 排程（每月 1 日） | 讀取月報數據，AI 分析趨勢，產出專業版/簡潔版/社群版/客戶版 4 種版本 | 11 |
| W4 | 節日關懷序列 | 排程（每天 07:30） | 自動偵測節日與客戶生日，AI 產出客製化關懷訊息 | 11 |
| W5 | 社群內容批量 | 排程（每月 1 日） | 批量產出 LINE 早安/FB 心得/IG 品牌/Threads 冷知識，存入排程表 | 14 |

## 匯入步驟

1. 開啟 n8n 介面（`http://localhost:5678` 或你的 n8n 雲端網址）
2. 點選左上角 **+** → **Import from File**
3. 選擇對應的 `.json` 檔案
4. 匯入後，設定每個 node 的 **Credentials**（API Key、OAuth 等）
5. 點選 **Activate** 啟用工作流

## 環境變數 / Credentials 設定

匯入後需要設定以下 Credentials：

| Credential 名稱 | 類型 | 用途 | 設定方式 |
|-----------------|------|------|----------|
| OpenAI API | API Key | AI 文案產出 | Settings → Credentials → 新增 OpenAI API → 貼上 API Key |
| Google Sheets OAuth | OAuth2 | 讀寫試算表 | Settings → Credentials → 新增 Google Sheets → OAuth2 授權 |
| LINE Notify Token | Header Auth | 發送 LINE 通知 | Settings → Credentials → 新增 Header Auth → `Authorization: Bearer YOUR_TOKEN` |

### 取得 Credentials

**OpenAI API Key**
- 前往 https://platform.openai.com/api-keys
- 建立新的 API Key，複製貼入 n8n

**Google Sheets OAuth**
- 前往 Google Cloud Console 建立 OAuth 2.0 Client
- 或在 n8n 中直接使用 Google OAuth 授權流程

**LINE Notify Token**
- 前往 https://notify-bot.line.me/my/
- 發行新的 Token，選擇要通知的群組

## Google Sheets 試算表結構

各工作流會讀寫以下試算表（匯入後需替換 Sheet ID）：

| 試算表名稱 | 工作表 | 用途 |
|-----------|--------|------|
| 物件資料庫 | `物件清單` | W1 存放行銷產出結果 |
| 客戶管理表 | `帶看紀錄` | W2 讀取客戶跟進排程 |
| 月報數據 | `月報` | W3 讀取業績數據 |
| 客戶名單 | `客戶資料` | W4 讀取生日/偏好 |
| 社群排程表 | `內容排程` | W5 存放批量產出的內容 |

## 注意事項

1. **API 費用**：每個工作流會呼叫 OpenAI API，請留意用量。建議使用 `gpt-4o-mini` 控制成本
2. **LINE Notify 限制**：每小時 1000 則，每則上限 1000 字
3. **Google Sheets ID**：匯入後需手動替換每個 Google Sheets node 中的 `spreadsheetId`
4. **時區設定**：排程觸發請確認 n8n 時區為 `Asia/Taipei`
5. **測試建議**：先用 **Execute Workflow** 手動測試，確認無誤後再 Activate
6. **備份**：啟用前建議先匯出原有工作流做備份

## 課程資訊

- 課程名稱：AI 超級房仲實戰班
- 講師：阿峰老師（黃敬峰）
- 官網：https://www.autolab.cloud
- 聯絡：ai@autolab.cloud
