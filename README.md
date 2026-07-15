# Rojnamcha Ledger - Shiv Shakti HP Gas Agency

A modern, responsive Web Application designed for managing the daily financial ledger, reports, cash audit reconciliation, and user access controls for **Shiv Shakti HP Gas Agency**.

This project is built using a decoupled architecture:
* **Frontend:** Hosted on Netlify as a single-page standalone web app shell.
* **Backend API:** Google Apps Script Web App.
* **Database:** Google Sheets (Spreadsheet).

---

## 🚀 Key Features

1. **Daily Entry Dashboard:**
   * Enter daily receipts and expenses.
   * Auto-calculates Net Balance, Cash Send to Home, Cash Deposit to Bank, and Closing Cash in Hand.
   * Tracks multiple cash transfers with detailed notes.

2. **Reports Dashboard:**
   * Filter financial records by custom date ranges.
   * Visual summaries of opening cash, total inflows/outflows, and final balances.
   * Export reports directly as **PDF** or **Email PDF** reports.

3. **Cash Audit & Slips:**
   * Denomination-based cash reconciliation (from ₹500 down to ₹1 coins).
   * Generates formatted audit summaries and slips.

4. **Dropdown Options Manager:**
   * Customize receipt/expense headers and set default rates for quick entries.

5. **User Access Controls:**
   * User registration and Admin approval flow.
   * Password hashing (SHA-256) for secure login.
   * **Admin Role** gains access to User Management (approvals, role edits, password resets).

6. **Database Backup & Restore (JSON):**
   * **Export Backup:** Downloads your entire database (Rojnamcha data, details, dropdowns, and users) as a single `.json` file.
   * **Import Backup:** Restores your entire database from a JSON backup file (requires confirmation keyword `RESTORE`).

---

## 🛠️ Setup Guide

### Step 1: Google Sheet Database Setup
1. Create a new Google Spreadsheet in Google Drive.
2. The script will automatically initialize the required sheets (`Rojnamcha`, `RojnamchaDetails`, `DropdownOptions`, `CompanyProfile`, `Users`) on first run.

### Step 2: Google Apps Script Setup
1. Open your Google Spreadsheet and go to **Extensions** -> **Apps Script**.
2. Rename the default `Code.gs` file to `Code.js` (or keep it as `Code.gs`) and paste the contents of `Code.js` from this repository.
3. Create the following HTML files in the Apps Script editor and copy-paste their respective contents:
   * `Index.html`
   * `JS.html`
   * `CSS.html`
   * `EmailTemplate.html`
4. Copy the contents of `appsscript.json` into the `appsscript.json` file in the Apps Script editor (turn on "Show appsscript.json manifest file" in editor settings if not visible).

### Step 3: Deploy Backend API
1. In the Apps Script editor, click **Deploy** (top-right) -> **New deployment**.
2. Select **Web app** as the deployment type.
3. Configure the deployment:
   * **Execute as:** `Me (your-email@gmail.com)`
   * **Who has access:** `Anyone` (required so the Netlify app can access the API).
4. Click **Deploy** and authorize the script permissions.
5. Copy the generated **Web App URL** (e.g. `https://script.google.com/macros/s/AKfy.../exec`).

### Step 4: Netlify Deployment
1. Log in to your Netlify account.
2. Click **Add new site** -> **Import from an existing project** and link your GitHub repository.
3. Configure the build settings:
   * **Build command:** (Leave blank)
   * **Publish directory:** `dist` (This serves the compiled static standalone web app).
4. Click **Deploy site**.

### Step 5: Connecting Frontend to Backend
1. Open your Netlify site URL (e.g., `https://rojnamcha.netlify.app/`).
2. If this is a fresh setup, you will be prompted with the **API Configuration Screen**.
3. Paste the **Google Web App URL** you copied in Step 3.
4. Log in using the default admin credentials:
   * **Username:** `admin`
   * **Password:** `admin`
5. Go to the **Manage Dropdowns** tab and scroll to the bottom to verify the **JSON Backup & Restore** utility is visible and working.

---

## 📦 Developer Guide (Clasp CLI)
This project uses `@google/clasp` to sync local files with Google Drive.

* **Install Clasp:** `npm install -g @google/clasp`
* **Login to Google:** `clasp login`
* **Pull latest server code:** `clasp pull`
* **Push local code to server:** `clasp push -f`
* **Compile Netlify file:** Run the Node compilation command to bundle `Index.html`, `CSS.html`, and `JS.html` into `dist/index.html`.
