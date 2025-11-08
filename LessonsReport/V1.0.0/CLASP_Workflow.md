# CLASP Workflow Guide

This guide explains how to manage your Google Apps Script project using CLASP (Command Line Apps Script Projects) with Git version control.

---

## 📁 Local Directory

Example project path:  
`C:\Projects\Maqamat\LessonsReport\V1.0.0`

---

## 🚀 1. Prerequisites

- [Node.js & npm](https://nodejs.org/) installed
- Google account with access to the script project
- [Git](https://git-scm.com/downloads) installed
- Run in Command Prompt (cmd) or Git Bash

---

## 🔐 2. Authenticate with Google

Open terminal and run:

bash
clasp login
## 📥 3. Clone (Download) from Apps Script to Local
If you already have the script open in Google Apps Script:

Go to Project Settings in Apps Script

Copy the Script ID

Then in terminal:

bash/cmd
Copy code to: C:\Projects\Maqamat\LessonsReport\V1.0.0
cd C:\Projects\Maqamat\LessonsReport\V1.0.0
clasp clone <SCRIPT_ID>
📁 This will download the script files into the current folder.

## 💾 4. Update Your Local Files
Edit any .gs or .html files locally using your preferred editor.

## 🆙 5. Push Changes Back to Google Apps Script
bash
Copy code
clasp push
🔁 Optional: Pull Latest From Google
bash
Copy code
clasp pull

## 🧹 6. Keep Local Files Out of GitHub
To keep this workflow file private, add it to .gitignore:

bash
Copy code
echo CLASP_Workflow.md >> .gitignore
Now Git will ignore this file when committing or pushing.

📝 Author
Yaniv Raba
📧 yaniv.raba@gmail.com