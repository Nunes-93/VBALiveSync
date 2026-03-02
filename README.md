# VBALiveSync: The Bridge for AI-Powered VBA Development 🚀

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python: 3.x](https://img.shields.io/badge/Python-3.x-blue.svg)](https://www.python.org/)
[![AI-Ready](https://img.shields.io/badge/AI%20Agent-Ready-purple.svg)]()

**VBALiveSync** is a revolutionary framework designed to enable AI Coding Agents (like Cursor, Windsurf, or GitHub Copilot) to develop, edit, and maintain VBA code directly within active Excel workbooks.

By using Python and the Windows COM interface, this tool bypasses the limitations of the native VBE, allowing a seamless, live, and safe two-way synchronization between modern IDEs and Excel.

## 🤖 The Core Innovation: AI-Driven Workflow

The true power of **VBALiveSync** resides in the `VBALiveSync.md` configuration. When provided to an AI Agent, it guides the AI to:

* **Autonomously Sync:** Pull code from Excel to a local environment, modify it, and push it back live.
* **Safe Operations:** Operate under strict "Zero-Trust" and "File Protection" rules, ensuring no binary corruption.
* **Modern Standards:** Use Git for versioning and modern IDE features while the AI handles the complex ASCII sanitization required for VBA stability.

## 🌟 Key Features

1. **AI Autonomy:** Provides the rules and tools for AI agents to write and inject macros safely.
2. **Live Two-Way Sync:** Code in your preferred IDE and see results in Excel instantly.
3. **Zero-Trust Architecture:** Mandatory environment checks to ensure the workflow is ready on any machine.
4. **Sanitization Engine:** Automatic refactoring of accents and special characters into `ChrW()` to prevent encoding errors.

## 🛠️ Prerequisites & Installation

* **OS:** Windows
* **Software:** Microsoft Excel installed and running
* **Libraries:** `pip install pywin32`

Clone the repository to your machine:

```bash
git clone [https://github.com/Nunes-93/VBALiveSync.git](https://github.com/Nunes-93/VBALiveSync.git)
cd VBALiveSync
```

🚀 How to Use with AI
Simply provide the VBALiveSync.md file as a System Prompt or Custom Instruction in your AI IDE. The Agent will immediately understand the architecture and begin managing your VBA modules through the vba_sync_auto.py script.

👤 Author
Nunes-93 - Modernizing legacy workflows through AI and automation.

💼 GitHub: @Nunes-93
