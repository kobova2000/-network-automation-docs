# Network Automation Tools  
### by Frank Abraham  

This repository contains two complementary scripts designed to automate the retrieval and documentation of Cisco switch configurations. The workflow is divided into two distinct phases:  

---

## **Overview**
1. **Phase 1 – Data Collection (Python)**
   - Script: `get-switch-configs.py`
   - Purpose: Connects to Cisco switches using **Netmiko**, retrieves the running configuration from each device, and saves each as an individual `.txt` file.
   - Output: Plain-text configuration files stored in the `final-configs` directory.

2. **Phase 2 – Documentation & Formatting (PowerShell)**
   - Script: `Build-Final-Configs-Doc.ps1`
   - Purpose: Reads all the collected `.txt` configurations and compiles them into a **professionally formatted Word document** for IT review, presentation, or compliance records.
   - Output: A finalized `final-configs.docx` report with clean headers, consistent spacing, and one configuration per page.

---

## **Project Workflow**
```text
Switch Devices  →  get-switch-configs.py  →  final-configs\*.txt
                                       ↓
                          Build-Final-Configs-Doc.ps1
                                       ↓
                        final-configs.docx (Word Report)

Requirements
Python Phase

Python 3.9+
Netmiko (pip install netmiko)
Windows or Linux environment
PowerShell Phase
Microsoft Word (required for COM automation)
Windows environment
PowerShell 5.1 or later

Usage
1. Run the Python script to collect configs

python get-switch-configs.py

Requires two local files:
-switch-targets.txt – list of switch IP addresses (one per line)
-creds.ini – credential file formatted as:
[auth]
username = your_username
password = your_password
secret = optional_enable_secret

2. Run the PowerShell script to generate the final document
.\Build-Final-Configs-Doc.ps1 -OutputDoc "C:\Path\To\final-configs.docx"

-Reads all .txt files in the final-configs folder.
-Compiles them into a single, formatted .docx report.

Features

-Secure multi-device SSH collection via Netmiko
-Automatically saves each configuration as a separate file
-Word automation with page breaks, headers, and consistent formatting

Ideal for network documentation, sign-off packages, or compliance deliverables

Sample Output

final-configs.docx
-Page 1: Hostname + IP header
-Body: Full configuration text
-Page 2–N: One device per page

### **Phase 2 – Documentation & Formatting (PowerShell)**  
**Script:** `Build-Final-Configs-Doc.ps1`  

This PowerShell script compiles all Cisco configuration text files into a single, professionally formatted Microsoft Word document.  
It is the final stage of the automation workflow that begins with `get-switch-configs.py`.

#### **Key Features**
- 30 spaces between hostname and IP for consistent alignment  
- Removes all lines containing `!`  
- Single-line spacing throughout (no extra spacing before or after)  
- Automatically inserts page breaks between configurations  
- Word runs silently in the background (no GUI)  
- Console-only logging for clean execution  

#### **Usage**
1. Verify that all configuration text files from Phase 1 are stored in:  
   `C:\Path\To\Your\Scripts\final-configs`
2. Run the script in PowerShell:
   ```powershell
   .\Build-Final-Configs-Doc.ps1 -OutputDoc "C:\Path\To\Your\Scripts\final-configs.docx"


License

Licensed under the MIT License
© 2025 Frank Abraham

Author

Frank Abraham
AI Integration & Automation Consultant
LinkedIn: linkedin.com/in/frankmabraham

