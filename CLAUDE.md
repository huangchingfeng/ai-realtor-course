# AI 超級房仲實戰班

## 專案概述
房仲 AI 實戰課程的完整教材與課後加值系統。5 模組 × 10 案例 = 50 個 Prompt 指令。

## 技術棧
- **PPT 生成**：Python + python-pptx
- **Prompt 指令庫**：純前端（HTML/CSS/JS + JSON）
- **n8n 工作流**：5 份 JSON 設計稿
- **Notion 模板**：Markdown 規格文件

## 目錄結構
```
├── index.html              # 課程介紹頁
├── prompt-library/         # Prompt 指令庫網站
│   ├── index.html
│   ├── css/style.css
│   ├── js/app.js
│   └── data/prompts.json
├── cases/                  # PPT 生成腳本 + 範例資料
│   ├── generate_ppt_module1~5.py
│   ├── extract_prompts.py
│   └── sample-data/
└── templates/              # 課後模板
    ├── n8n/               # 5 個工作流 JSON
    └── notion/            # Notion 規格文件
```

## 開發指令
```bash
# 重新提取 Prompt 資料
cd cases && python3 extract_prompts.py

# 本地測試 Prompt 指令庫
cd prompt-library && python3 -m http.server 8888

# 生成 PPT
cd cases && python3 generate_ppt_module1.py
```

## 品牌色
- 主色：#00D4FF（Cyan）
- 背景：#0A1628（Navy）
- 強調：#FF6B35（Orange）
- 文字：#FFFFFF / #C7C7CC

## 密碼
- Prompt 指令庫：`realtor2026`
