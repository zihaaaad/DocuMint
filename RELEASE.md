# Release Notes - v3.0.0

**Release Date:** September 2026  
**Version:** 3.0.0 (Major Release - Production Standard)

## 🚀 What's New in v3.0.0

### 1. Unified Multi-Syntax Template Engine
* Full support for `<Column>`, `{{Column}}`, `[Column]`, and `{Column}` variable tokens.
* Complete traversal across Word document paragraphs, multi-level tables, and section headers & footers (`section.header` / `section.footer`).
* Run-level character formatting preservation (bold, italic, font family, RGB color).

### 2. Universal PDF Converter Cascade & Session Optimization
* Persistent Word COM automation with automatic instance reuse and clean lifecycle management (`DispatchEx` + `CoInitialize`).
* Multi-tier fallback cascade: Windows Word COM &rarr; `docx2pdf` &rarr; Headless LibreOffice on Linux/macOS/Docker.

### 3. Asynchronous Multi-Channel Dispatch Engine
* Multi-threaded asynchronous dispatch via `ThreadPoolExecutor` (5 concurrent workers) for SMTP.
* Automatic SSL (Port 465) vs. STARTTLS (Port 587/25) negotiation in `SMTPSender`.
* Native COM integration for Microsoft Outlook desktop client.

### 4. ACID SQLite Audit Ledger & Analytics
* SQLite WAL mode (`history.db`) for concurrent audit logging without database locks.
* Granular per-recipient execution tracking and historical success rate analytics.

### 5. Headless CLI Interface
* Brand new command-line parser (`documint run`, `documint validate`, `documint studio`, `documint stats`).
* Enables automated batch execution in headless Linux servers, CI/CD pipelines, and scheduled tasks.

### 6. Modern Web Studio Dashboard
* Real-time progress monitoring, live email previewer, and System Logs console.
* Interactive Profile Management (Save, Load, and Delete preset JSON profiles).
* Dry-Run mode and Single Test Recipient execution support.
* Live job cancellation and abort signal handling.

---
*Authored and maintained by [Zihad Hasan](https://zihadhasan.web.app)*
