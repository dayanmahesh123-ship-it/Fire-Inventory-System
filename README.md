# 🧯 Fire Protection Inventory Management System

A professional, cloud-based inventory tracking solution for fire safety equipment. This system is designed to manage stock levels for items like fire extinguishers, smoke detectors, and safety gear using a high-tech dashboard.

## 🚀 Tech Stack
* **Backend:** Google Apps Script (GAS) [cite: 2026-03-03]
* **Database:** Google Sheets (Acts as a real-time database) [cite: 2026-03-01]
* **Frontend:** React 18, Tailwind CSS, and Babel Standalone (All via CDN) [cite: 2026-03-05]

## ✨ Key Features
* **Real-time Synchronization:** Data is fetched and updated directly from Google Sheets without refreshing the page. [cite: 2026-03-03]
* **Inventory Management:** Full CRUD functionality to add, view, update (sell), and delete equipment records. [cite: 2026-03-05]
* **Low Stock Alerts:** Automatic visual indicators highlight any item with a stock count below 5 in red. [cite: 2026-03-03]
* **Responsive Design:** A clean, dark-themed professional UI that works on both desktop and mobile devices. [cite: 2026-03-05]

## 🛠️ Setup Instructions for Beginners
1. **Spreadsheet:** Create a Google Sheet and name the first tab `Inventory`. Add headers: `Item Name`, `Price`, `Stock`. [cite: 2026-03-01]
2. **Apps Script:** Go to `Extensions > Apps Script` in your Google Sheet. [cite: 2026-03-03]
3. **Files:** Copy the `Code.gs` and `index.html` from this repository into the script editor. [cite: 2026-03-03]
4. **Deploy:** Click `Deploy > New Deployment`, select `Web App`, set access to `Anyone`, and click Deploy. [cite: 2026-03-03]



## 📁 Project Structure
* `Code.gs` - Handles the server-side logic and spreadsheet communication. [cite: 2026-03-03]
* `index.html` - Contains the React 18 frontend, Tailwind styles, and Babel compiler. [cite: 2026-03-05]
* `README.md` - Documentation and project overview. [cite: 2026-03-05]

---
**Developed by Mahesh** - *Learning and building modern web systems with Google Cloud and React.* [cite: 2026-03-03]
