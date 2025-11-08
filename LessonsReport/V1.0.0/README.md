# 📋 דיווח שיעורים - Google Apps Script Project

This project automates lesson reporting (`דיווח שיעורים`) in a Google Sheet.  
It manages lesson data, enforces editing rules, updates statuses, and logs every run — all through custom menus, triggers, and sidebars.

---

## 📁 Project Files Overview

| File | Purpose |
|------|----------|
| **Sidebar.html** | Main sidebar UI for running lesson updates |
| **BulkStatusSidebar.html** | Sidebar for quick monthly status updates |
| **00_Utils.gs** | General helper utilities for dates, text, etc. |
| **01_Bootstrap.gs** | Initialization, triggers, and UI menu setup |
| **02_Main.gs** | Orchestrator: main workflow controller |
| **05_LogSvc.gs** | Logging service (writes to “לוג ריצות”) |
| **10_SheetsSvc.gs** | Sheet creation, formatting, and utilities |
| **11_RowsSvc.gs** | Handles row-level logic and duplication prevention |
| **12_ExceptionsSvc.gs** | Manages exception filtering for group lessons |
| **13_ProtectSvc.gs** | Handles row locking/unlocking and background colors |
| **20_GroupSvc.gs** | Logic for processing group lessons |
| **21_PrivateSvc.gs** | Logic for processing private lessons |
| **30_PostProc.gs** | Post-processing (sorting, coloring, counters) |
| **40_BulkStatusUpdate.gs** | User-facing bulk update handler |
| **41_BulkStatusQuickUpdate.gs** | Fast back-end monthly status update |
| **70_HighlightSvc.gs** | Highlights cells with issues |
| **99_README.gs** *(optional)* | Inline summary if README.md not synced |

---

## ⚙️ Trigger Configuration

| Setting | Value |
|----------|--------|
| **Function:** | `onEditShowLock` |
| **Event Source:** | From spreadsheet |
| **Event Type:** | On edit |
| **Notifications:** | Notify me daily |

This trigger automatically detects attempts to edit locked rows and shows a custom alert or dialog.

---

## 🚀 Features

- ✅ Auto-lock rows when status = “שולם” or “הועבר לתשלום”
- 🔓 Auto‑unlock rows when status = “דווח‑טרם שולם”
- 🧮 Calculates “סך השיעורים שנותרו”
- 📊 Writes logs to the “לוג ריצות” sheet
- 🟡 Highlights missing or invalid data
- 🗓 Supports monthly bulk status updates
- 🧩 Handles exception filters for group lessons
- 🧰 Fully modular service‑based architecture

---

## 🛠️ Setup Instructions

### 1️⃣ Prepare Google Sheets
Make sure your spreadsheet contains the following sheets:
- `דיווח שיעורים`
- `רשימת קורסים-מערכת`
- `ריכוז שיעורים פרטיים`
- `חריגים-קבוצתי`
- `סטטוס`
- `לוג ריצות`

### 2️⃣ Install the Apps Script Code
Copy all `.gs` and `.html` files into your Apps Script editor  
or manage them using **clasp** (see below).

### 3️⃣ Configure the Trigger
- Open Apps Script → Triggers (⏱ icon)
- Create a new trigger:
  - Function: `onEditShowLock`
  - Source: From spreadsheet
  - Type: On edit
  - Notifications: Notify me daily

### 4️⃣ Run Once
Run the `onOpen()` function once manually to register menus and permissions.

---

## 🧠 Developer Notes

- The project uses `setWarningOnly(false)` to **fully block editing** on locked rows.
- Gray background = locked row  
  Yellow background = needs attention  
- Status, Lock, and Log tracking are centralized under `LogSvc` and `ProtectSvc`.

---

## 👤 Author

**Yaniv Raba**  
📧 yaniv.raba@gmail.com

---

