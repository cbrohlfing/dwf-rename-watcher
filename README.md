# DWF Rename Watcher

A lightweight background tray application that automatically renames DWF files based on matching IDW filenames.

Built with PowerShell + WinForms for engineering automation workflows.

---

## 🚀 Overview

DWF Rename Watcher monitors a folder for newly created `.dwf` files and automatically renames them to match their corresponding `.idw` file.

This eliminates manual renaming steps and ensures consistent file naming between drawing formats.

The application runs quietly in the Windows system tray.

---

## 🔧 What It Does

When a new `.dwf` file appears in the watched directory:

1. The script checks for a matching `.idw` file
2. If a match is found, the `.dwf` file is renamed to match the `.idw` filename
3. Logging is written for traceability
4. The app continues monitoring in the background

---

## 📂 Features

- Real-time folder monitoring
- Runs minimized in the system tray
- Auto-start capable
- Configurable settings via JSON
- Safe renaming logic with error handling
- Debug logging support
- Versioned release structure

---

## ⚙ Configuration

Settings are stored in: dwf_rename_settings.json

Typical configurable values include:

- Watch directory path
- Logging options
- Rename behavior rules

---

## 🖥 Running the Application

Launch using:

Rename-IdwDwf-Watcher-Tray.ps1

Or via the provided desktop shortcut.

The application will appear in the Windows system tray.

Right-click the tray icon to exit.

---

## 🏷 Versioning

Releases follow a `vX.Y` format.

Example:

- v1.3
- v1.4
- v1.5

Older legacy `vX.Y.Z` versions are still detected correctly.

---

## 🛠 Tech Stack

- PowerShell
- .NET FileSystemWatcher
- WinForms (Tray UI)
- JSON configuration
- Git version control

---

## 🎯 Intended Use

Designed for engineering environments where:

- IDW files generate DWF outputs
- Naming consistency is critical
- Manual renaming wastes time
- Automation improves reliability

---

## 🔒 Notes

- Designed for Windows environments
- Requires appropriate file system permissions
- Works on local or network paths

---

## 👨‍💻 Author

Chris Rohlfing
Engineering Automation Projects
