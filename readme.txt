# WIN_INSPECTION_SCRIPT

**Windows System Inspection & Auditing Tool (Handover Documentation)**

This document provides an explanation of the WIN_INSPECTION_SCRIPT project's code structure, features, setup instructions, and maintenance information. It is intended to help understand, modify, and extend the script collection as needed.

## 1. Project Overview

WIN_INSPECTION_SCRIPT is a collection of PowerShell-based scripts designed for collecting information and auditing Windows systems. It features a GUI interface allowing users to select specific workstations and perform tasks such as checking basic information, viewing audit history, and backing up event logs.

## 2. Features

* **Graphical User Interface (GUI):** Provides an interactive way to manage inspection tasks via `MainPanel.ps1`.
* **Workstation Selection:** Easily select the target workstation using a dropdown menu (`$computerComboBox`).
* **Basic Information Logging:** Record and retrieve simple notes or basic information for each workstation, saved to `Report\[Workstation Name]\basic_info.txt`.
* **Audit History Browse:** View a list (`$auditHistoryListView`) of timestamped audit records stored for each workstation (`Report\[Workstation Name]\[Timestamp]`). Displays folder timestamp, compliance info (if applicable), and logged-in user.
* **HTML Report Generation:** Generate detailed HTML reports (`audit_report.html`) from selected audit data (includes account info, service info from `.json` files, and security logs from `.log` file) via a right-click context menu.
* **Event Log Backup:** Launch a dedicated GUI (`backupScript.ps1`) to select and back up specific Windows event logs (e.g., System, Security, Application) to timestamped folders (`log\[Workstation Name]\[Timestamp]`). Includes quick selection buttons for major/all logs.
* **Customizable Background:** Allows setting a custom background image for the main GUI.

## 3. File Structure

The project's main file and folder structure is as follows:
WIN_INSPECTION_SCRIPT/
├── script/
│   ├── MainPanel.ps1         # Main GUI interface and logic
│   ├── InspectionModule.ps1  # (Currently unused) Module for system inspection functions (future expansion)
│   └── backupScript.ps1      # Event log backup GUI and logic
├── Report/                 # Root folder for workstation reports and info
│   └── [Workstation Name]/   # Folder specific to a workstation
│       ├── [YYYYMMDD_HHmmss]/ # Timestamped audit record folder
│       │   ├── account.json      # Example: Account information
│       │   ├── service.json      # Example: Service information
│       │   └── security_report.log # Example: Security log content
│       └── basic_info.txt      # Basic notes/info for the workstation
├── etc/                    # Miscellaneous files/resources (e.g., background images)
│   └── background/
│       └── [background_image_file] # Background image(s) for the GUI
├── log/                    # Root folder for event log backups (created by backupScript.ps1)
│   └── [Workstation Name]/   # Workstation specific backup folder
│       └── [YYYYMMDD_HHmmss]/ # Timestamped backup folder containing event logs (.evtx files)
└── README.md               # This documentation file (originally readme.txt)

*(Note: The `log/` directory structure is inferred from the description of `backupScript.ps1`)*

## 4. Key Script Descriptions

### 4.1. `script/MainPanel.ps1`
* **Role:** Provides the main GUI interface and handles user interactions and core logic for the inspection tool.
* **Key Functions:**
    * **Workstation Selection (`$computerComboBox`):** Allows the user to select the target workstation from a dropdown list.
    * **Basic Info (`$basicInfoTextArea`, `$basicInfoChangeButton`):** Loads and saves basic textual information/notes related to the selected workstation into `Report\[Workstation Name]\basic_info.txt`.
    * **Audit History (`$auditHistoryListView`):** Displays a list of timestamped audit folders found within `Report\[Workstation Name]\`. Shows the folder name (timestamp), compliance information, and logged-in user associated with each audit record.
    * **View Report (Right-Click Menu):** When an item in the Audit History list is right-clicked, this option generates an `audit_report.html` file by combining information from `account.json`, `service.json`, and `security_report.log` within the selected audit folder, then opens the report in the default web browser.
    * **Log Backup Button (`$logBackupButton`):** Executes the `script\backupScript.ps1` script, passing the currently selected workstation name as a parameter.
    * **Background Image:** Sets the background of the main GUI using an image file located in the `etc/background/` folder.

### 4.2. `script/backupScript.ps1`
* **Role:** Provides a dedicated GUI and the underlying logic specifically for backing up Windows event logs.
* **Key Functions:**
    * **Log Selection (`$logCheckBoxList`):** Presents a list of available event logs, allowing the user to select which ones to back up via checkboxes.
    * **Quick Select Buttons:** Provides buttons to quickly select commonly needed logs ("Major Logs") or all available logs ("All Logs").
    * **Start Backup (`$startButton`):** Initiates the event log backup process for the logs selected by the user.
    * **Backup Structure:** Saves the exported event logs (typically as `.evtx` files) into a structured folder path: `log\[Workstation Name]\[YYYYMMDD_HHmmss]\`.

### 4.3. `script/InspectionModule.ps1`
* **Role:** (Currently unused) This module was intended to house functions and logic related to system inspection tasks, such as gathering detailed system information or checking specific configurations. It remains available for future expansion of the tool's capabilities.

## 5. Setup and Execution

1.  **Verify File Structure:** Ensure the project files and folders are organized correctly as described in Section 3. You may need to manually create the base `Report/` and `log/` directories if they do not exist.
2.  **Background Image:** To customize the GUI background, place your desired image file (e.g., `.jpg`, `.png`) into the `etc/background/` folder. Then, update the filename reference within the `MainPanel.ps1` script to point to your new image file.
3.  **Execute Script:** Open PowerShell and navigate to the project's root directory. Run the `MainPanel.ps1` script.

    ```powershell
    # Example assuming you are in the WIN_INSPECTION_SCRIPT directory
    .\script\MainPanel.ps1
    ```
    The main GUI window should appear.

## 6. Prerequisites

* **PowerShell:** Requires a compatible version of Windows PowerShell. (It's advisable to specify the minimum tested version, e.g., PowerShell 5.1 or higher).
* **Permissions:** Running these scripts, particularly functions related to system information gathering and event log backup, will likely require elevated privileges. Run PowerShell **as Administrator**.

## 7. Maintenance and Extension

* **Coding Style:** Maintain code consistency by adhering to standard PowerShell coding practices and conventions.
* **Error Handling:** Utilize `try-catch` blocks effectively throughout the scripts to handle potential errors gracefully and provide meaningful feedback to the user.
* **Logging:** Consider implementing logging mechanisms to record script execution details, errors, or significant events. This can greatly assist in troubleshooting.
* **Modularization:** For adding significant new features, encapsulate related functions within separate script files or PowerShell modules (`.psm1`) to improve organization and reusability. Referencing the planned `InspectionModule.ps1` is a good starting point.
* **Future Expansion Ideas:**
    * Implement various system information gathering and configuration check functions within `InspectionModule.ps1`.
    * Offer additional report export formats (e.g., CSV, JSON) besides HTML.
    * Include more comprehensive data points in the generated reports.
    * Add capabilities to inspect remote workstations (requires handling PowerShell Remoting, authentication, etc.).
    * Introduce functionality for scheduling automated audit runs.

## 8. Precautions

* **Permissions:** Always ensure that the environment where the scripts are run has the necessary permissions (likely Administrator rights) for the intended operations.
* **Testing:** Before deploying any modifications or enhancements, conduct thorough testing in a controlled environment to prevent unexpected behavior or data loss.

