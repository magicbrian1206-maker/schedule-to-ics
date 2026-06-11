# 📅 排班 → 行事曆轉換器

把排班 Excel 一鍵轉成可匯入 Apple/Google 行事曆的 `.ics` 檔。

🔒 **隱私**：100% 純前端，Excel 檔不會離開你的瀏覽器。

## 線上使用
👉 https://YOUR-USERNAME.github.io/schedule-to-ics/

## Excel 格式
標題列需有：`Eng Name`（人員）、`Date`（日期）、`Status`（假別代碼）

支援代碼：`AL` `SL` `M` `MT` `PT` `B` `BT` `CL` `NP` `OFF` `RPH` `PH1`–`PH16`、Apple Holiday

## 自動規則
- 加班偵測：當天有人排 `PH1–PH16`，其他人若狀態是 `All day in store` 自動標為「加班」
- 行事曆標題格式：`{First Name} {Code}`（例：`Brian AL`）
- 全部為整日事件，TRANSPARENT（不擋忙碌時段）
- `OFF` 公休不匯出（網頁上會顯示）

## 本機執行
直接雙擊 `index.html` 即可。
