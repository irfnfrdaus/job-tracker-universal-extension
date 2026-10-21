# 🎯 JobPulse — Universal Job Application Tracker

[![Manifest V3](https://img.shields.io/badge/Chrome%20Extension-Manifest%20V3-blue)](https://developer.chrome.com/docs/extensions/mv3/)
[![Cross Platform](https://img.shields.io/badge/OS-Windows%20%7C%20macOS%20%7C%20Linux-brightgreen)](#-optional-desktop-excel-sync)
[![Browsers](https://img.shields.io/badge/Browsers-Chrome%20%7C%20Edge%20%7C%20Opera%20%7C%20Brave%20%7C%20Arc-orange)](#-browser-installation-60-seconds)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

**JobPulse** is an open-source, universal browser extension and application cockpit designed to eliminate messy job application spreadsheets. 

Log applications from **LinkedIn, JobStreet, Indeed, Glassdoor, and career portals** in one click. Works **100% standalone** right inside any browser, with an **optional real-time desktop Excel sync** for power users on Windows, macOS, and Linux.

---

## ✨ Key Features

- **⚡ 1-Click Smart Auto-Detect:** Automatically grabs Job Title, Company Name, Platform, and URL directly from the active tab.
- **🌐 Full In-Browser Cockpit:** Includes a built-in dark mode dashboard with search, filtering, and metric cards (`Total`, `Applied`, `Interviewing`, `Offers`, `Rejected`).
- **📥 CSV & Excel Export:** Export your entire job hunt database to CSV in one click for Google Sheets, Numbers, or Excel.
- **💻 Desktop Excel Sync (Optional):** Automatically writes formatted rows with colored status tags to your local `.xlsx` file via a lightweight Python bridge.
- **✏️ Live Status Updater:** Change application statuses (e.g., from `Applied` to `Interviewing` or `Offer Received`) directly from the in-browser dashboard or live web view with 1 click, syncing instantly to your desktop Excel spreadsheet.
- **🛡️ Unblockable File-Lock Queue:** If Microsoft Excel is open while you save, jobs are queued safely without throwing file-lock (`PermissionError`) errors and flushed automatically the second Excel is closed.
- **🔄 Seamless Offline Sync:** Jobs saved while the bridge is disconnected automatically sync and receive sequential IDs the moment the bridge comes online.
- **🔒 100% Private & Local:** Zero external cloud servers. Your data stays entirely in your browser storage or on your own hard drive.

---

## 🚀 Browser Installation (60 Seconds)

Works on **Google Chrome, Opera, Opera GX, Microsoft Edge, Brave, and Arc**.

### 1. Download or Clone the Repository
```bash
git clone https://github.com/irfnfrdaus/job-tracker-universal-extension.git
```
*(Or click **Code ➔ Download ZIP** on GitHub and extract it to a folder on your computer).*

### 2. Load into Your Browser

| Browser | Steps |
| :--- | :--- |
| **Google Chrome / Brave** | Go to `chrome://extensions` ➔ Toggle **Developer mode** (top right) ➔ Click **Load unpacked** ➔ Select the `extension/` folder. |
| **Opera / Opera GX** | Go to `opera://extensions` ➔ Toggle **Developer mode** ➔ Click **Load unpacked** ➔ Select the `extension/` folder. |
| **Microsoft Edge** | Go to `edge://extensions` ➔ Toggle **Developer mode** (bottom left) ➔ Click **Load unpacked** ➔ Select the `extension/` folder. |
| **Arc Browser** | Go to `arc://extensions` ➔ Enable **Developer mode** ➔ Click **Load unpacked** ➔ Select `extension/`. |

That's it! Pin **JobPulse** to your toolbar and start applying!

---

## ⚡ Mode 1: Standalone (No Setup, No Python Needed)

JobPulse works completely out of the box with zero external dependencies:
1. Browse to any job opening on **LinkedIn**, **JobStreet**, or **Company Portals**.
2. Click the **JobPulse** extension icon.
3. Review auto-detected fields and click **Save Application**.
4. Click **🌐 Open Dashboard** to see all your applications, filter by status, search, and export to CSV anytime!

---

## 📊 Mode 2: Optional Desktop Excel Sync (Windows, Mac, Linux)

If you love having an active Microsoft Excel `.xlsx` file on your desktop that auto-updates every time you click "Save", follow these simple steps:

### Requirements
- Python 3.8+ installed ([python.org](https://www.python.org/downloads/))

### Windows Setup:
1. Open the `bridge/` folder.
2. Double-click **`start_bridge.bat`**.
3. It installs `openpyxl` and creates `Job_Applications_Tracker.xlsx` inside your `Documents/Job_Tracker/` folder automatically!

### macOS / Linux Setup:
1. Open Terminal and navigate to the `bridge/` directory:
   ```bash
   cd bridge
   bash start_bridge.sh
   ```
2. Your master spreadsheet is created at `~/Documents/Job_Tracker/Job_Applications_Tracker.xlsx`.

Whenever the bridge is running, the extension status changes to **"Excel Connected"**, and every application is written directly into your local Excel spreadsheet with custom styles and borders.

---

## 📁 Repository Structure

```
job-tracker-universal-extension/
├── extension/                  # Universal Browser Extension (Manifest V3)
│   ├── manifest.json           # Permissions & browser config
│   ├── popup.html              # 1-Click logging popup
│   ├── popup.js                # Auto-scraping & save logic
│   ├── dashboard.html          # Built-in full-screen analytics dashboard
│   ├── dashboard.js            # Search, filter, and CSV export logic
│   └── icons/                  # 16px, 48px, 128px high-res icons
├── bridge/                     # Optional Cross-Platform Python Bridge
│   ├── bridge.py               # Auto-detects OS & creates master .xlsx
│   ├── requirements.txt        # Python dependencies (openpyxl)
│   ├── start_bridge.bat        # 1-Click Windows launcher
│   └── start_bridge.sh         # 1-Click macOS/Linux launcher
├── .gitignore
├── LICENSE                     # MIT License
└── README.md
```

---

## 🤝 Supported Platforms for Auto-Scraping

- [x] **LinkedIn Jobs**
- [x] **JobStreet**
- [x] **Indeed**
- [x] **Glassdoor**
- [x] **Handshake**
- [x] **MyFutureJobs**
- [x] **Direct Company Career Sites** (Workday, Greenhouse, Lever, SmartRecruiters, Taleo)

---

## 📄 License

Distributed under the **MIT License**. Free for students, job-seekers, and developers worldwide. See `LICENSE` for details.

---

**Developed with ❤️ by [Muhammad Irfan Firdaus](https://github.com/irfnfrdaus)**  
*UTP Big Data Analytics | Ex-IT QA Automation Intern at Hilti Asia IT Services*
