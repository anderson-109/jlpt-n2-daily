# 📚 JLPT N2 每日學習表

每天早上 **10:00（台灣時間）** 自動生成一份 JLPT N2 單字＋語法 Word 學習表，並上傳到本 Repo。

## 📁 檔案結構

```
├── .github/
│   └── workflows/
│       └── daily_n2.yml      # GitHub Actions 排程設定
├── scripts/
│   └── generate_n2.py        # Word 文件生成腳本
├── docs/
│   └── JLPT_N2_YYYY-MM-DD.docx  # 每日生成的學習表
└── README.md
```

## 📄 學習表內容

每天從題庫中隨機抽取：
- **3 組 N2 單字**：單字、讀音、詞性、中文意思、例句（含振假名）
- **3 組 N2 語法**：語法型式、用法說明、接續方式、例句（含振假名）

## 🔧 手動觸發

在 GitHub 頁面：**Actions → 每日 JLPT N2 學習表生成 → Run workflow**

## 📥 下載學習表

到 `docs/` 資料夾選擇對應日期的 `.docx` 檔案下載即可。
