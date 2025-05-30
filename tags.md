## 🏷️ Tags - Windows Server 2012 R2 Hardening (CIS v2.2.1)

This repository includes a legacy audit script for Windows Server 2012 R2, based on the CIS Benchmark v2.2.1. Below is a categorized list of concepts and techniques applied.

---

### 🛡️ Security Hardening Topics

- CIS Benchmark alignment
- Domain Controller vs. Member Server distinction
- Registry value inspection
- WMI queries

---

### 🔐 Authentication and Account Policies

- Password complexity
- Maximum password age
- Minimum password age
- Account lockout threshold
- Lockout duration
- Lockout observation window

---

### ⚙️ Service and Protocol Restrictions

- Disabling unnecessary services (e.g., Telnet, Remote Registry)
- Disabling LM hashes
- Restrict anonymous access
- Disable SMBv1

---

### 🗂️ System Settings and Registry Keys

- Audit Policy settings
- User Rights Assignment
- Security Options in Local Policies
- Event log configuration (size, retention)

---

### 📋 Auditing and Logging

- Audit account logon events
- Audit object access
- Audit policy change
- Audit privilege use
- Audit system events

---

### 🛠️ Script Features

- Written in VBScript (VBS)
- Read-only audit (no system modifications)
- Output printed to standard output (no logging to files)
- Compatible with cscript.exe
- Manual execution (no scheduler, no automation)

---

### 🧾 Legacy Notice

- Not compatible with Windows Server 2016 or newer
- No PowerShell version available
- Script is archived and no longer maintained
