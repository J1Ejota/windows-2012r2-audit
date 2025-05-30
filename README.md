# 🧾 Windows Server 2012 R2 Hardening Audit (Legacy)

This repository contains a **VBScript-based audit tool** developed in 2017 to assess Windows Server 2012 R2 configurations against the **CIS Benchmark v2.2.1**.

> ⚠️ **Legacy Notice**: This script is **deprecated** and intended for **educational or historical reference only**. It targets **Windows Server 2012 R2** and is **not compatible** with modern versions (e.g., 2016, 2019, 2022).

---

## 📌 About This Script

- **Script name**: `Windows-benchmark.vbs`
- **Language**: VBScript (VBS)
- **Target system**: Windows Server 2012 R2
- **Benchmark**: CIS Microsoft Windows Server 2012 R2 Benchmark v2.2.1
- **Execution**:

```cmd
cscript.exe //nologo Windows-benchmark.vbs
```

The script performs **read-only** audits against various CIS controls. It distinguishes between **Domain Controllers** and **Member Servers** using WMI and registry queries.

---

## 📂 Repository Structure

```plaintext
windows-2012r2-audit/
├── Windows-benchmark.vbs     # The main audit script (legacy)
├── docs/
│   └── CIS_Windows_2012R2_v2.2.1.pdf   # Reference benchmark (for offline use)
├── LICENSE
├── README.md
├── changelog.md
└── .gitignore
```

---

## 🚫 Limitations

- No longer maintained
- Written in VBScript, not PowerShell
- Not compatible with Windows Server 2016 or newer
- Does not generate reports or structured outputs (e.g., CSV)

---

## 📄 License & Credits

Created by **José Manuel Campuzano**  
[LinkedIn](https://www.linkedin.com/in/jose-manuel-campuzano) · [GitHub](https://github.com/J1Ejota) · [HackTheBox](https://app.hackthebox.com/profile/984522)

---

## ✅ Recommendation

If you're managing modern Windows Server environments, consider using:

- **PowerShell** scripts
- **Group Policy Management** with ADMX templates
- **Windows Security Baselines** from Microsoft
- **Modern CIS Benchmarks** for Windows Server 2019/2022

This script remains archived here as a technical reference and a personal milestone.
