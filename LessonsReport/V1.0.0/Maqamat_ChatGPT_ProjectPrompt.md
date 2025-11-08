🔁 Prompt for Future ChatGPT Sessions
You’re assisting me with a Google Apps Script project titled:

**"Maqamat – LessonsReport"**

This project is a fully automated spreadsheet-based system used to manage, update, and lock lesson reporting data for a music school (Maqamat). The system supports monthly report generation, bulk status updates, access control via row locking, and visual UI feedback (toasts, dialogs, background colors, icons). It is written in **Google Apps Script** with HTML sidebar integration.

---

# Maqamat – LessonsReport

This is a Google Apps Script project for managing and automating monthly lesson reporting for the Maqamat music school. The system integrates spreadsheet logic with UI elements, automated row locking, and advanced logging and filtering.

---

## 🔁 Project Summary

- **Project Title:** Maqamat – LessonsReport  
- **Platform:** Google Apps Script (with HTML Sidebar)  
- **Author:** Yaniv Raba (yaniv.raba@gmail.com)  
- **Assistant (AI Agent):** ChatGPT (OpenAI)  

---

## 🔧 Key Functionalities

1. **Main Report Flow (`App.processReport`)**  
   - Handles parsing, importing, and inserting group/private lessons.  
   - Applies exception filters and avoids duplicates.  
   - Locks rows with finalized statuses.

2. **Bulk Status Updates**  
   - Allows monthly status update (e.g. "שולם - אסור לערוך שינויים").  
   - UI feedback through color and status message.

3. **Validation and Locking**  
   - Enforces dropdowns for status values.  
   - Locks rows where changes should be disallowed.  
   - Only newly inserted rows are considered for locking.

4. **Visual & UI Enhancements**  
   - Highlighting (yellow for flagged issues).  
   - Custom message boxes and toast notifications.

5. **Logging & Tracking**  
   - All actions are logged into:
     - `סטטוס` sheet — run status summary.
     - `לוג ריצות` sheet — detailed log entries per run.
   - Structured logs per step and run ID.

---

## 📂 Files Included



---

### 📂 File List (Script Editor Files):
- `00_Utils.gs`
- `01_Bootstrap.gs`
- `02_Main.gs`
- `05_LogSvc.gs`
- `10_SheetsSvc.gs`
- `11_RowsSvc.gs`
- `12_ExceptionsSvc.gs`
- `13_ProtectSvc.gs`
- `20_GroupSvc.gs`
- `21_PrivateSvc.gs`
- `30_PostProc.gs`
- `40_BulkStatusUpdate.gs`
- `41_BulkStatusQuickUpdate.gs`
- `70_HighlightSvc.gs`
- `Sidebar.html`
- `BulkStatusSidebar.html`

---

### ⚙️ Trigger Configuration:
- **Trigger Function:** `onEditShowLock`
- **Event Type:** `On edit`
- **Source:** Spreadsheet
- **Deployment Version:** Head
- **Failure Notifications:** Notify me daily

> This trigger monitors edits on the “דיווח שיעורים” sheet, and when a locked row is edited it displays a custom modal dialog to the user.

---

## 🧰 Local Development (CLASP)

- Local directory path: `C:\Projects\Maqamat\LessonsReport\V1.0.0`  
- Code synced with Google Apps Script via `clasp`.  
- Project versioned via GitHub; `.gitignore` includes:  
---

### 🧰 Dev Workflow:
- Project managed via `clasp` on local machine.
- Stored under: `C:\Projects\Maqamat\LessonsReport\V1.0.0`

so that the local workflow documentation is not pushed to the repo.

---

## 🧠 Assistant Notes

This entire Apps Script project is managed in collaboration between you and **ChatGPT (OpenAI)**.  
Feel free to reference this prompt in future sessions so I can quickly recall the project’s context, structure, and next steps.

## 🧠 AI Agent
Apps Script: Expert consultant for Google Apps Script coding. 
By maksymstoianov.com