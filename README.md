# Office Online Repair Script

![GitHub license](https://img.shields.io/badge/license-MIT-blue.svg)

A PowerShell script to silently perform an **online repair** for Microsoft Office Click-to-Run installations. It automatically logs its actions and prevents conflicts with ongoing repairs.

---

## Features

* **Silent Online Repair:** Automates Office repair without user interaction.
* **Intelligent Checks:** Verifies Office installation and ongoing repairs.
* **Detailed Logging:** Creates unique log files in `C:\Windows\Temp\` for each run.

---

## Prerequisites

* **Windows OS** with **Microsoft Office Click-to-Run**.
* **PowerShell 5.1+**.
* **Administrative privileges** are required to run the script.

---

## Installation & Usage

1.  **Download** the `O365OnlineRepair.ps1` script from this repository.
2.  **Open PowerShell as Administrator.**
3.  **Navigate** to the script's directory: `cd C:\Path\To\Your\Script`
4.  **Unblock** the script (if downloaded from the internet): `Unblock-File -Path .\OfficeOnlineRepair.ps1`
5.  **Run** the script: `.\O365OnlineRepair.ps1`

### **Deployment with IT Management Tools**

This script is ideal for deployment using tools like **PDQ Deploy**, **Microsoft Endpoint Configuration Manager (SCCM)**, **Intune**, **NinjaOne**, or other RMM (Remote Monitoring and Management) solutions.

You can typically:
* Create a new package or application in your chosen deployment tool.
* Add a "PowerShell Script" step and upload this `O365OnlineRepair.ps1` file.
* Configure the deployment to run with administrator privileges on target machines.

---

## Logging & Exit Codes

* **Logs** are saved as `OfficeRepair_YourComputerName_YYYYMMDD_HHmmss.log` in `C:\Windows\Temp\`.
* **Exit Code 0:** Script completed successfully, or exited gracefully (Office not found/repair already running).

---

## Troubleshooting

* **"Access Denied"**: Run PowerShell as **Administrator**.
* **Script doesn't run**: Check the log file in `C:\Windows\Temp\` for details.

---

## License

This project is licensed under the [MIT License](LICENSE).
