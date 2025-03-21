# PEPATO
# 📦 PrePac Tool - PrePackaging Utility for Windows Software

**PerPac Tool** is a powerful PowerShell-based GUI application to assist IT professionals and software packagers in preparing and analyzing Windows software installations. This tool provides snapshot functionality, MSI analysis, and installer observation for creating clean and documented software packages.

---

## ✨ Features

- 🔥 **Firewall Rule Snapshots**
  - Take "Before" and "After" snapshots
  - Compare differences in rules

- 🧠 **Registry Snapshots**
  - Capture "Before" and "After" snapshots of HKLM\Software
  - Compare and export changes

- 📂 **Installer Execution & Tracking**
  - Run EXE installers as SYSTEM
  - Manually observe extracted MSI files

- 🧪 **MSI Analysis**
  - Extract `ProductCode`, `REBOOT`, `ALLUSERS`
  - List all shortcuts and their target locations

- 🕵️ **Human-Friendly Interface**
  - Built with WPF in PowerShell
  - Central log output
  - Easy-to-navigate button layout

---

## 📁 Output Directory

All snapshot files are stored in:
```
C:\temp\PrePackTool
```
Each file is timestamped using the format:
```
dd_MM_yy
```
Example:
```
FirewallSnapshot_Before_21_03_25.txt
```

---

## 🚀 How to Run

1. Make sure you run PowerShell **as Administrator**
2. Ensure execution policy allows scripts:
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```
3. Run the script:
   ```powershell
   .\PerPacTool.ps1
   ```

---

## 🛠 Requirements

- Windows 10/11
- PowerShell 5.1+
- Admin rights (required for firewall and registry export)

---

## 📌 Roadmap Ideas

- ✅ Export logs to HTML or Markdown
- ⏳ Progress bar support during snapshots
- 📊 Deeper MSI table inspection (CustomActions, Features)
- 🧼 Clean-up automation after snapshot

---

## 📃 License

MIT License. Use at your own risk. Contributions welcome!

---

## 🤝 Contributing

Pull requests, ideas, and improvements are welcome! Let’s simplify software packaging together 🙌

---

Made with ❤️ by packagers, for packagers.

