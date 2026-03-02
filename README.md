# VBALiveSync: Modern & AI-Ready VBA Development 🚀

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python: 3.x](https://img.shields.io/badge/Python-3.x-blue.svg)](https://www.python.org/)
[![AI-Ready](https://img.shields.io/badge/AI%20Agent-Ready-purple.svg)]()

**VBALiveSync** is a revolutionary open-source workflow that bridges the gap between modern IDEs, AI coding agents, and legacy VBA (Visual Basic for Applications).

Say goodbye to the native VBA Editor (VBE) locked inside binary Excel files. With this tool, you can finally use Git for version control, code in a modern dark theme, and let AI agents write and inject macros safely into your live Excel applications.

## 🌟 The Revolution

Traditionally, VBA is trapped in binary files. External manipulation usually corrupts the file. **VBALiveSync** solves this by implementing a **Zero-Trust, Live Two-Way Sync architecture:**

1. **Safe Extraction (Pull):** Pulls code directly from an *open* Excel workbook into clean, local text files (`.bas`, `.cls`, `.frm`).
2. **Modern Editing:** Edit your VBA in VS Code or your preferred modern editor.
3. **Live Injection (Push):** Push the code back into the live Excel file instantly using Python's `win32com` interface. The physical binary file is never touched directly, preventing corruption.
4. **ASCII Sanitization:** Built-in rules to handle character encoding.

## 🤖 The Killer Feature: AI Agent Integration

This repository is a complete framework designed for AI-assisted development. By using the included `setup_vba.md` file as a system prompt in your AI IDE, you give the AI the exact boundaries it needs to safely develop VBA code for you autonomously.

## 🛠️ Prerequisites & Installation

* **OS:** Windows
* **Software:** Microsoft Excel installed and running.
* **Libraries:** `pip install pywin32`

Clone the repository to your machine:
```bash
git clone https://github.com/Nunes-93/VBALiveSync.git
cd VBALiveSync
