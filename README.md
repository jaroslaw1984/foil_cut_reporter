# Foil Cut Reporter

A production support application designed to generate reports on logistics processes related to the cutting of laminating film. 

## 📷 Screenshots

### Main Application Dashboard
![Main Window](assets/main_window.png)


### A new report has been released
![Print View](assets/main_window_ready_to_print.png)

## 🚀 Key Features

- **Automated Data Refresh:** Continuous monitoring of the availability of new reports for a given machine. The program checks every 30 seconds to see if a new     document has appeared on the server's hard drive.
- **Generating a report:** The report is generated based on the data contained in the JSON file; it then combines all the data, after which the report is created as a Word file. 
- **Auto-Update Notification System:** Built-in version checking mechanism that automatically triggers a pop-up window notifying the user whenever a newer release is deployed on the network server.

## 🛠️ Tech Stack

- **Language:** Python 3.10+
- **GUI Framework:** CustomTkinter
- **Data Processing:** pandas, openpyxl
- **Distribution:** PyInstaller (standalone `.exe` for Windows 11)

## 📁 Project Structure

```text
foil_cut_reporter/
│
├── assets/                  # Application graphical assets and resources
├── config/
│   ├── paper_foils.json     # Configuration file for foils data
│   ├── paths.py             # Global path management definitions
│   └── version.py           # Application version definition
│
├── database/
│   └── db_manager.py        # Database connection and query manager
│
├── deploy/
│   ├── deploy_gui.py        # Deployment scripts for UI modules
│   ├── deploy_logic.py      # Deployment core logic configuration
│   └── paths.py             # Deployment specific paths
│
├── gui/
│   ├── components/
│   │   ├── machine_card.py  # Reusable UI component representing a single machine
│   │   └── popup_about.py   # Information popup modal
│   └── app_view.py          # Main application window framework
│
├── logic/
│   └── report_engine.py     # Core foil cutting optimization and reporting logic
│
├── updater/
│   ├── icon.ico             # Executable icon asset
│   ├── icon.png             # Application logo image
│   ├── PyInstaller_HowTo.txt # Internal guidelines for compiling the app
│   └── updater.py           # Remote version checking and update logic
│
├── .gitignore               # Git untracked files configuration
├── app.py                   # Application main entry point
├── README.md                # This documentation file
├── requirements.txt         # Project third-party dependencies
└── run_deploy.py            # Script triggering deployment and distribution builds
