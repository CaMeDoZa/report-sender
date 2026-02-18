# Report Sender 📧

Automated email report distribution system - Evolution from production PowerShell

## 🎯 Project Goal

Transform a production Windows automation script into a cloud-native application while learning DevOps practices.

## 📂 Project Structure
```
report-sender/
├── powershell/          # Original Windows implementation
│   └── report-sender.ps1
├── bash/                # Linux Bash rewrite
│   ├── file_waiter.sh
│   ├── config.sh
│   └── mailer.sh
├── python/              # Python implementation (coming soon)
├── docs/
└── README.md
```

## 🚀 Current Phase: Bash Implementation

### Original Requirements:

**Business Logic:**
- Monitor directories for daily report files (Excel)
- Wait up to 20 minutes for files to appear
- Send reports via SMTP with attachments
- Support multiple recipients with CC
- Generate execution summaries

**Technical Features:**
- Retry logic (5 attempts, 45s delay)
- Comprehensive logging (INFO, SUCCESS, ERROR, CRITICAL)
- File lock detection
- Timeout handling
- Exit codes for automation

### Bash Features (in progress):

- [ ] File monitoring with timeout
- [ ] Structured logging system
- [ ] Configuration management
- [ ] Modular architecture
- [ ] Error handling & exit codes

## 🛠️ Tech Stack

**Current (Bash):**
- Shell: Bash 5.x
- OS: Debian 12

**Future (Python + Docker):**
- Language: Python 3.11+
- Container: Alpine Linux
- Orchestration: Kubernetes

## 📖 Original Implementation

The PowerShell script automated daily report distribution:
- File monitoring for multiple regions
- SMTP email sending (Mail.ru)
- HTML email templates
- Retry mechanism with logging
- Daily summary reports

**Impact:**
- Automated distribution for 10+ regions
- Saved ~2 hours/day of manual work
- 99% delivery success rate

## 🧑‍💻 Author

**Dan** ([@CaMeDoZa](https://github.com/CaMeDoZa)), telegram: @camedoza