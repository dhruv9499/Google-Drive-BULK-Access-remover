# 🧹 Google Drive Email Access Remover

**Because manually removing ex-employees/Freelancers from 500+ files is not how I want to spend my Friday evening.**

[![Made with Google Apps Script](https://img.shields.io/badge/Made%20with-Google%20Apps%20Script-4285f4.svg)](https://script.google.com/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](http://makeapullrequest.com/)

A powerful Google Apps Script tool that automatically removes specified email addresses from **all Google Drive files** accessible by your account. Perfect for lazy people (like me) who don't want to manually clean up file permissions when team members leave.

## 🎯 The Problem

When freelancers, employees, or associates leave your organization, their email addresses often remain in file sharing permissions across hundreds (or thousands) of Drive files. Manually removing them is:

* ⏰ **Time-consuming** - Who has time to check every single file?
* 😴 **Boring** - Repetitive clicking is mind-numbing
* 🐛 **Error-prone** - Easy to miss files or make mistakes
* 😤 **Frustrating** - There's no bulk "remove user from everything" button

## ✨ The Solution

This script does the boring work for you:

```javascript
// Just add the emails you want to remove
const TARGET_EMAILS = [
  'former.employee@company.com',
  'freelancer@example.com'
];

// Run once and forget
startEmailCleanup();
```

## 🚀 Features

* 🎯 **Smart Email Search** - Uses Drive API queries to find files shared with specific emails
* 📧 **Bulk Processing** - Remove multiple email addresses in one run
* 🛡️ **Self-Protection** - Won't let you accidentally remove yourself
* ⏱️ **Execution Time Management** - Handles Google Apps Script's 6-minute limit automatically
* 📊 **Detailed Reporting** - Get a comprehensive email summary when complete
* 🔄 **Auto-Resume** - Continues processing across multiple script runs
* 🚨 **Robust Error Handling** - Gracefully handles permission issues
* 💤 **Set & Forget** - Start it and go grab coffee

## 🏃‍♀️ Quick Start

### 1. Copy the Script

**[📋 Click here to copy the Google Apps Script project](https://script.google.com/d/156cHRxfungPRmYWCPSJI8JwK3y3nzwXg3ToyR7Wz8uj7fyVtSp60Y6BS/edit?usp=sharing)**

*(Just click "Make a copy" and you're ready to go!)*

### 2. Enable Drive API

* In your copied project: **Services** → **Add a service** → **Drive API**

### 3. Configure Target Emails

```javascript
const TARGET_EMAILS = [
  'person1@company.com',
  'person2@freelancer.com'
  // Add as many as you need
];
```

### 4. Test & Run

```javascript
// Test everything works
runDiagnostics();

// Start the cleanup
startEmailCleanup();
```

### 5. Get Coffee ☕

The script handles everything automatically and emails you a detailed report when done.

## 📊 What You Get

### Real-time Progress Monitoring

```javascript
checkStatus(); 
// 📊 Process Status:
//    📧 Current email: john@company.com (2/3)
//    📁 Files processed: 127
//    ⏰ Started: 2/14/2025, 2:30:00 PM
//    ⏳ Running time: 8 minutes
```

### Comprehensive Email Report

* **By Email** : Files found, successful removals, errors, success rate
* **By File Type** : Google Docs, Sheets, Slides, PDFs, etc.
* **Detailed File Lists** : Exactly which files were processed
* **Error Details** : What went wrong and why

## 🧪 Testing Functions

Before running on all your files, test everything works:

| Function                | What It Does                                                        |
| ----------------------- | ------------------------------------------------------------------- |
| `runDiagnostics()`    | **🩺 Full system check**- Tests everything                    |
| `testDriveV3Access()` | Tests Drive API connection                                          |
| `testEmailSearch()`   | Tests searching for files with target emails                        |
| `checkStatus()`       | Shows current progress                                              |
| `stopProcess()`       | Emergency stop button                                               |
| `testForEdgeCases()`  | Tests for files with permission limitations (limited effectiveness) |

## ⚙️ How It Works

### The Smart Approach

Instead of scanning every single file in your Drive, this script:

1. **Searches Directly** - Uses `'email@domain.com' in readers or writers or owners` queries
2. **Processes by Email** - Handles one target email at a time
3. **Batches Intelligently** - Processes files in small groups to avoid timeouts
4. **Auto-Continues** - Resumes where it left off if interrupted

### The Technical Magic

* **Drive API v3** - Latest and fastest Google Drive API
* **Smart Pagination** - Handles large datasets efficiently
* **State Management** - Remembers progress between script runs
* **Error Recovery** - Continues despite permission issues
* **Resource Optimization** - Works within Apps Script limits

## 🔧 Configuration Options

```javascript
const CONFIG = {
  BATCH_SIZE: 15,                    // Files per batch (adjust for speed)
  MAX_EXECUTION_TIME: 280000,        // Stop before 6-minute limit
  RETRY_DELAY: 5000,                 // Seconds between batches  
  MAX_LOG_SIZE: 8000                 // Log size before truncation
};
```

## 🚨 Safety Features

* ✅ **Self-Protection** - Can't remove your own email
* ✅ **Email Validation** - Checks email format before processing
* ✅ **Permission Handling** - Skips files you can't access
* ✅ **State Recovery** - Handles crashes gracefully
* ✅ **Comprehensive Logging** - Full audit trail of all actions

## 🐛 Troubleshooting

### "No files found" Issue

```javascript
// Run full diagnostics first
runDiagnostics();

// This will tell you exactly what's wrong:
// ✅ Drive API v3: Working
// ✅ Emails Configured: 3 emails found  
// ❌ Email Search: No files found
```

### Common Solutions

* **Drive API Not Enabled** : Add "Drive API" from Services
* **No Target Emails** : Add emails to `TARGET_EMAILS` array
* **Permission Issues** : You can only remove access from files you own/manage
* **Process Stuck** : Use `stopProcess()` and restart

### 🔒 Known Limitation: Files with Limited Access

The script **cannot detect or modify** files where:

1. You have only Viewer or Commenter access
2. The target email has Editor or higher access
3. Another user owns the file

This is due to a limitation in Google Drive's API that restricts permission searches based on your access level. Even though you can see these files in the Drive UI, the API does not return them in search queries.

**Solution for these edge cases:**

- Have the file owner run the script instead
- Manually remove access through the Google Drive web interface
- Use admin tools if you have domain administrator privileges

## 📈 Performance

* **Small Organizations** (~100-500 files): 5-15 minutes
* **Medium Organizations** (~1000-5000 files): 30-60 minutes
* **Large Organizations** (5000+ files): 1-3 hours

*Performance depends on number of files, API quotas, and how many contain target emails.*

## 🤝 Contributing

Found a bug? Have an idea? PRs welcome!

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Test thoroughly with your own Drive files first
4. Commit changes (`git commit -m 'Add AmazingFeature'`)
5. Push to branch (`git push origin feature/AmazingFeature`)
6. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](https://claude.ai/chat/LICENSE) file for details.

## ⚠️ Disclaimer

**Test in a safe environment first!**

* Always make a backup copy of your script
* Test with non-critical files initially
* Understand that bulk permission changes can't be easily undone
* This tool only works on files you own or have permission to manage

## 🙏 Acknowledgments

* Built with Google Apps Script and Drive API v3
* Inspired by lazy developers who hate repetitive tasks
* Tested by frustrated managers cleaning up after departed team members

---

## 🎉 Ready to Reclaim Your Friday Evening?

1. **[📋 Copy the script](https://script.google.com/d/156cHRxfungPRmYWCPSJI8JwK3y3nzwXg3ToyR7Wz8uj7fyVtSp60Y6BS/edit?usp=sharing)**
2. Enable Drive API in Services
3. Add your target emails
4. Run `startEmailCleanup()`
5. Go do something more interesting!

**Questions? Issues?** Open an issue or check the troubleshooting section above.


# **Sample Email Image**


---

*Made with ❤️ and excessive amounts of caffeine by a developer who got tired of manually removing people from Google Drive files.*
