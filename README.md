# Outlook Microsoft 365 Reply Tracker
A VBA + Outlook automation tool that tracks email responses and sends daily summary reports.

工具：
Excel VBA (Macro) + Outlook COM + Windows Task Scheduler

做什麼事：
1. 每日掃描收件匣中特定寄件者的通知信
2. 自動判別每封信是否有回覆
3. 將結果寫入 Excel 並於固定時間寄出彙總清單

價值：
取代人工逐一檢查 Outlook 的作業，節省每天 15-30 分鐘，確保不漏追回覆，並自動產出報表留存。

### Email Output Example

```
Subject: [Auto Notification] Daily Summary - June

本月通知清單：

Department      Date        Replied   Reply Time           Subject
Trading Desk    2026/06/04  Yes       2026/06/05 08:06     Trading Desk Alert Notice
Trading Desk    2026/06/03  Yes       2026/06/03 19:00     Trading Desk Alert Notice
Risk Management 2026/06/02  No                             Risk Management Alert Notice
```

## Project Structure

```
outlook-notification-tracker/
├── vba/                          # VBA modules
│   ├── ModuleFullProcess_DEMO.bas     # Full pipeline (fetch + check replies + send email)
├── scripts/
│   └── RunDaily.vbs              # Task Scheduler trigger script
├── excel/                        # Excel workbook template
├── .gitignore
└── README.md
```

## How It Works

1. Scans the Outlook Inbox for alert emails from `alert@company.com` with **current month** only
2. Extracts: department name, date, subject, body summary
3. Searches for **reply** emails (same date + keyword, different sender)
4. Writes results to Excel (replied → `Yes` + timestamp, not replied → `No`)

## Requirements

- Windows 10/11
- Microsoft Excel (macros enabled)
- Microsoft Outlook

## Setup

1. Create ReplyTracking.xlsm, import vba/ModuleFullProcess_DEMO.bas via VBA editor (Alt+F11 > File > Import)
2. Schedule via Task Scheduler: daily at 09:00, run wscript.exe "C:\Path\RunDaily.vbs"

## Usage
```
Macro	               Description
RunAll	             Fetch alerts + check replies, write to Excel
RunAllAndSendEmail	 RunAll + email summary
```
```
VBA logic excerpt (company project, full code not disclosed)

For i = inboxFolder.Items.Count To 1 Step -1
    If TypeOf item Is Outlook.MailItem Then
        If sender = "alert@company.com" And InStr(subject, "ANC") Then
            If FindReply(dateStr, "ANC") Then
                ' Mark as replied
            End If
        End If
    End If
Next
```
