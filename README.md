# Ultimate Email Automation Suite
*A powerful, all-in-one email automation and management platform*

## 📋 Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Installation](#installation)
- [Prerequisites](#prerequisites)
- [Configuration](#configuration)
- [Usage](#usage)
- [Options](#options)
- [Troubleshooting](#troubleshooting)
- [Important Notes](#important-notes)
- [Disclaimer](#disclaimer)

## 🚀 Overview

The **Ultimate Email Automation Suite** is a comprehensive Python-based tool designed for professional email management, collection, and campaign automation. It provides an intuitive interface for scraping emails from websites, managing email campaigns, and organizing your email outreach efforts.

## ✨ Features

### 🔍 Email Collection
- **Web Scraping**: Automatically extract emails from multiple websites
- **Contact Page Detection**: Intelligent navigation to contact pages
- **Batch Processing**: Process multiple URLs simultaneously
- **Duplicate Removal**: Automatic deduplication of collected emails

### 📧 Email Campaign Management
- **Bulk Email Sending**: Send personalized emails to multiple recipients
- **Attachment Support**: Include CVs, demand letters, and other documents
- **Custom Templates**: Use pre-written email templates
- **Progress Tracking**: Real-time sending progress with success/failure tracking

### 🔧 Advanced Email Management
- **Gmail Integration**: Search and manage sent emails directly
- **Email Search**: Find emails by subject in your sent folder
- **Email Deletion**: Manage and clean up your sent folder
- **Data Export**: Save collected emails to organized files

### 🎯 User Experience
- **Beautiful Interface**: Color-coded, professional terminal interface
- **Progress Indicators**: Visual progress bars and status updates
- **Session Statistics**: Detailed analytics and performance metrics
- **Error Handling**: Comprehensive error handling with helpful messages

## 📥 Installation

### Step 1: Install Python
Ensure you have **Python 3.7 or higher** installed:
```bash
python --version
```

### Step 2: Install Required Dependencies
```bash
pip install pyppeteer imaplib smtplib email pyfiglet termcolor getpass
```

### Step 3: Install Browser Dependencies
The script uses Microsoft Edge for web scraping. Ensure you have:
- **Microsoft Edge** browser installed
- Edge WebDriver (usually comes with Edge installation)

### Step 4: Download the Script
Save the provided Python script as `email_automation.py`

## 🔧 Prerequisites

### Gmail Account Configuration

#### 1. Enable 2-Factor Authentication (2FA) - REQUIRED
**📺 Watch this video tutorial: [How to Enable 2FA on Gmail](https://youtu.be/kTcmbZqNiGw?si=Sh0OzSuXg2-zqVH-)**
- **Watch from: 0:00 to 2:19**
- This is mandatory for app password generation
- Follow the steps exactly as shown in the video

#### 2. Generate App Password
After enabling 2FA:
1. Go to your Google Account settings
2. Navigate to **Security** → **2-Step Verification** → **App passwords**
3. Generate a new app password for "Mail"
4. Save this 16-character password (you'll use it in the script)

#### 3. Enable IMAP Access - REQUIRED
**📺 Watch this video tutorial: [How to Enable IMAP in Gmail](https://youtu.be/LSdCoz_J3fc?si=t6f5og27AHtp3Y3R)**
- **Watch the entire video carefully**
- This enables email management features (search, delete, etc.)
- Without this, email management features won't work

#### Steps to Enable IMAP:
1. Open Gmail in your browser
2. Click the gear icon → **See all settings**
3. Go to **Forwarding and POP/IMAP** tab
4. Enable **IMAP access**
5. Click **Save Changes**

## ⚙️ Configuration

### File Structure Preparation
Create the following file structure:
```
project_folder/
├── email_automation.py
├── urls.txt          # List of websites to scrape
├── letter.txt        # Email template
├── cv.pdf           # Your CV (optional)
└── demand.pdf       # Demand letter (optional)
```

### Required Files Setup

#### 1. URLs File (`urls.txt`)
Create a text file with one URL per line:
```
https://company1.com
https://company2.com
https://company3.com/about
https://company4.com/contact
```

#### 2. Email Template (`letter.txt`)
First line = Subject, rest = Email body:
```
Application for Software Developer Position

Dear Hiring Manager,

[Your email content here...]

Best regards,
[Your Name]
```

#### 3. Attachments (Optional)
- `cv.pdf` - Your curriculum vitae
- `demand.pdf` - Your salary requirements or cover letter

## 🎮 Usage

### Running the Script
```bash
python email_automation.py
```

### Main Menu Options

#### 1. 📧 Collect Emails
- Extracts emails from websites listed in `urls.txt`
- Automatically navigates to contact pages
- Saves collected emails to file

#### 2. 🚀 Send Campaign
- Sends emails to collected addresses
- Uses your Gmail account
- Supports attachments
- Provides real-time progress tracking

#### 3. 📥 Import Emails
- Load emails from existing files
- Multiple format support (TXT, CSV, etc.)
- Options to add to current collection or replace

#### 4. 👁️ View Emails
- Browse collected email addresses
- Display statistics and counts

#### 5. 🔄 Clear Data
- Reset collected emails
- Requires confirmation to prevent accidental deletion

#### 6. 🔍 Email Management
- **Search Sent Emails**: Find emails by subject
- **Delete Emails**: Remove emails from sent folder
- **Export Emails**: Save found emails to file
- **Quick Send**: Send to found emails immediately

#### 7. 📊 Session Report
- View detailed statistics
- Success rates and performance metrics
- Session duration and activity summary

#### 8. ❌ Exit
- Graceful exit with session summary
- Clean termination of all processes

## 🔄 Complete Workflow

### Step-by-Step Process

1. **Setup Phase**
   ```
   Enable 2FA → Generate App Password → Enable IMAP → Prepare Files
   ```

2. **Collection Phase**
   ```
   URLs File → Web Scraping → Email Extraction → Data Saving
   ```

3. **Campaign Phase**
   ```
   Load Emails → Compose Email → Add Attachments → Send Campaign
   ```

4. **Management Phase**
   ```
   Search Sent Emails → Analyze Results → Export Data → Clean Up
   ```

## ⚠️ Important Notes

### Security Considerations
- 🔒 **Never share your app password**
- 🔒 **Use a dedicated Gmail account for automation**
- 🔒 **Keep your email templates professional**
- 🔒 **Respect email sending limits**

### Gmail Limits
- **Daily sending limit**: 500 emails per day
- **Rate limiting**: ~100 emails per hour recommended
- **Attachment size**: 25MB maximum

### Best Practices
1. **Test with small batches first**
2. **Personalize email content**
3. **Respect anti-spam regulations**
4. **Monitor success rates**
5. **Keep backups of your data**

## 🛠️ Troubleshooting

### Common Issues

#### ❌ "Login Failed" Error
- Verify 2FA is enabled
- Use app password (not your main password)
- Check IMAP is enabled

#### ❌ "No Emails Found"
- Check if websites have public email addresses
- Verify URLs are accessible
- Try manual navigation to contact pages

#### ❌ "Sending Failed"
- Check daily sending limits
- Verify recipient email formats
- Ensure attachments aren't too large

#### ❌ "Browser Launch Failed"
- Verify Microsoft Edge installation
- Check Edge WebDriver availability
- Run script as administrator if needed

### Performance Tips
- 🚀 Use high-quality URL lists
- 🚀 Split large campaigns into batches
- 🚀 Monitor Gmail sending limits
- 🚀 Keep attachments optimized

## 📞 Support

### Required Video Tutorials
Make sure to watch these essential tutorials:

1. **2-Factor Authentication Setup**
   - 📺 [Watch Here](https://youtu.be/kTcmbZqNiGw?si=Sh0OzSuXg2-zqVH-)
   - ⏱️ Watch from: 0:00 to 2:19

2. **IMAP Activation**
   - 📺 [Watch Here](https://youtu.be/LSdCoz_J3fc?si=t6f5og27AHtp3Y3R)
   - ⏱️ Watch entire video

## 📄 License

This project is for educational and professional use. Users are responsible for complying with:
- CAN-SPAM Act regulations
- GDPR requirements
- Local email marketing laws
- Terms of service of email providers

## ⚠️ Disclaimer

This tool is designed for legitimate email outreach and professional communication. Users must:
- Obtain proper consent where required
- Respect unsubscribe requests
- Comply with all applicable laws
- Use responsibly and ethically

---

**Made with ❤️ by Otmane Sniba**
