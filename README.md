# 台灣稅務 LLM Wiki

對外教學用的台灣稅務知識庫，第一階段涵蓋**綜合所得稅**與**房地合一所得稅**。參考 Karpathy 的 LLM Wiki 做法：原始法規存在 `raw/`，由 AI 依規範整理成 `wiki/` 知識頁，再做成教學簡報。

**簡報網站：** https://112303085-jpg.github.io/tax-wiki/ （只放「定稿」簡報）

## 這個專案怎麼確保正確

- 所有法條、金額、日期都來自 `raw/` 裡的官方原文（全國法規資料庫、財政部），不用 AI 的記憶補。
- 引用的條文由程式逐字比對原文；例題的每個數字都由程式實際計算。
- 每個 wiki 頁標明查證狀態；人工審閱過的頁面才能用在對外簡報。
- 每次推送，GitHub 會自動重跑全部檢查（見 Actions 分頁）。

## 資料夾

| 資料夾 | 內容 |
|---|---|
| `raw/` | 原始來源（法規、函令、官方說明），存入後不修改 |
| `wiki/` | 知識頁：稅目、條文、概念、數額表、來源摘要、爭點；`index.md` 是目錄 |
| `slides/` | 簡報，每份一個資料夾，含大綱、逐頁內容、檢查報告與成品（HTML、PowerPoint） |
| `tools/` | 檢查與產生工具：`law.py`、`calc.py`、`deck.py`、`build_site.py` |
| `assets/` | 簡報用的圖示與素材 |

工作規範寫在 [`AGENTS.md`](AGENTS.md)，Claude Code、Codex 等 AI 工具會自動讀取。

## 在本機執行檢查

需要 Python 3.12 以上（只用標準函式庫）：

```
python tools/law.py check wiki/條文/所得稅法第17條.md
python tools/calc.py check wiki/條文/所得稅法第17條.md
python tools/deck.py status slides/所得稅介紹
```

產生簡報成品（`deck.py render`）需要 Windows、Microsoft PowerPoint 與 Edge。

## 免責聲明

本專案內容為教學整理，僅供學習參考，不構成個案稅務或法律意見。法規與數額可能修正，實際適用請以財政部及主管機關最新公告為準，或洽詢專業人士。
