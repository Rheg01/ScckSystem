# SCCK ROTC Command Center & Management System 🛡️

Hey there! Welcome to the repository for the SCCK ROTC Command Center.

I built this full-stack web application to solve a real-world logistical nightmare: managing attendance, discipline records, and academic grades for an ROTC unit of over 300+ cadets. Relying on paper trails and disorganized spreadsheets was causing data loss and massive delays.

This system digitizes the entire workflow, introducing offline-first QR scanning, role-based access control, and a real-time command dashboard.

## 💻 The Tech Stack

I wanted to keep the architecture lightweight, serverless, and highly accessible without relying on expensive hosting.

* **Frontend:** HTML5, CSS3, Vanilla JavaScript (ES6+)
* **Backend / REST API:** Google Apps Script (Server-side JavaScript)
* **Database:** Google Sheets (Real-time data syncing)
* **Libraries:**
* `Chart.js` for real-time dashboard data visualization
* `qr-scanner` for optical QR code parsing
* `SheetJS` for exporting tables to perfectly formatted Excel files
* `SweetAlert2` for responsive UI modals and alerts



## 🚧 Challenges I Faced (and Overcame)

Building a live production system from scratch taught me a lot. Here were the biggest hurdles:

* **The "Offline Scanner" Problem:** ROTC training often happens in open fields with terrible internet connectivity. If the connection dropped, attendance scanning would fail. I engineered a failsafe using the browser's `LocalStorage`. If the system detects it's offline, it caches the QR scans locally. Once the connection is restored, the user can hit a "Sync Offline" button to push the entire array to the backend API in one batch.
* **Cache Desyncs on the Dashboard:** To make the dashboard load instantly, I used Google's `CacheService` to store the database payload. However, when an admin issued a Merit or Demerit, the dashboard wouldn't reflect the new score because it was reading the old cache. I had to build a dynamic cache-busting mechanism that surgically deletes the specific cache keys the moment a new discipline log is successfully saved to the database.
* **Stubborn Browser Layouts:** I ran into a massive headache where native browser date-pickers and dropdowns were stubbornly overflowing out of their CSS containers on smaller screens. I completely refactored the UI from horizontal `Flexbox` wrappers into a strict, vertically stacked `CSS Grid` architecture to mathematically force the elements to respect their parent boundaries.

## 🚀 How to Use / Setup

Because this app runs on a serverless Google Apps Script architecture, you don't need `npm install` or a local Node environment to get it running.

### 1. Database & Backend Setup

1. Create a new Google Sheet and set up the following tabs: `Accounts`, `Masterlist`, `Attendance`, `Discipline`, `Leaves`, `AuditLog`, `Legacy_Classes`, `Legacy_Officers`, `Public_Posts`, and `Command_Staff`.
2. Go to **Extensions > Apps Script** in your Google Sheet.
3. Paste the contents of `backend.js` (from this repository) into the script editor.
4. Click **Deploy > New Deployment** as a Web App (Set access to "Anyone").
5. Copy the generated Web App URL.

### 2. Frontend Connection

1. Open the `app.html` file in this repository.
2. Locate the `GAS_API_URL` variable at the top of the JavaScript section.
3. Paste your Web App URL into this variable.
4. Open `app.html` in any web browser. The system will auto-initialize, connect to your database, and create a default Super Admin account so you can log in!

---

> **Note:** This project was built as a modern, cost-effective solution for military/ROTC logistics. If you have any questions about the code or want to collaborate, feel free to reach out!
