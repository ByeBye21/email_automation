# Email Automation Application

A comprehensive email automation tool with SMTP authentication, CSV recipient management, and modern GUI interface.

## ✨ Features

- 📧 **SMTP Authentication** - Support for Gmail, Outlook, Yahoo, iCloud, and custom servers
- 📊 **Flexible CSV Management** - Smart recipient handling with any column structure
- ✏️ **Rich Email Composer** - Built-in editor with dynamic attribute insertion
- 📎 **File Attachments** - Support for multiple file attachments with size display
- 🚀 **Bulk Email Campaigns** - Progress tracking and real-time campaign management
- 📋 **Activity Logging** - Comprehensive logging with HTML export functionality
- 🎨 **Modern Interface** - Professional card-based responsive design
- ⚙️ **Configuration Management** - Save/load email settings and templates
- 🔄 **Auto-refresh** - Real-time CSV file monitoring and recipient updates


## 📦 Installation

### Prerequisites

- Python 3.7 or higher
- PyQt5


### Steps

1. **Clone or download the repository:**

```bash
git clone https://github.com/ByeBye21/email_automation.git
cd email_automation
```

2. **Install dependencies:**

```bash
pip install -r requirements.txt
```

3. **Run the application:**

```bash
python email_gui.py
```


## 📋 Requirements

```txt
PyQt5>=5.15.0
```


## 🚀 Quick Start

### 1. Setup Email Account

- **Choose Provider:** Gmail, Outlook, Yahoo, iCloud, or Custom
- **Enter Credentials:** Your email address and password/app password
- **SMTP Settings:** Auto-configured for major providers
- **Test Connection:** Verify your email settings work

**Important for Gmail users:** Use an App Password instead of your regular password:

1. Enable 2-Factor Authentication on your Google account
2. Generate an App Password: Google Account → Security → App Passwords
3. Use the generated 16-character password in the application

### 2. Load Recipients

- **Upload CSV:** Select your recipients file (default: `./recipients.csv`)
- **Column Selection:** Choose which column contains email addresses
- **Preview Data:** Review recipients and available attributes
- **No Format Restrictions:** Any CSV structure is supported


### 3. Compose Your Email

- **Subject Line:** Write your email subject (supports {{attributes}})
- **Email Content:** Use the rich text editor for your message
- **Personalization:** Click attribute dropdowns to insert {{name}}, {{company}}, etc.
- **Attachments:** Add files with drag-and-drop or browse button


### 4. Launch Campaign

- **Test First:** Send a test email to verify everything works
- **Monitor Progress:** Real-time progress tracking with statistics
- **Campaign Results:** Detailed success/failure reporting
- **Activity Logs:** Complete audit trail of all operations


## 📝 CSV Format

Your CSV file can have any structure. Common examples:

**Basic Format:**

```csv
email,name,company
john.doe@example.com,John Doe,Acme Corp
jane.smith@example.com,Jane Smith,Tech Solutions
```

**Extended Format:**

```csv
email_address,first_name,last_name,company_name,position,city
john.doe@example.com,John,Doe,Acme Corporation,Software Engineer,New York
jane.smith@example.com,Jane,Smith,Tech Solutions,Marketing Manager,Los Angeles
```

**Any Column Names Work:**

```csv
contact_email,full_name,business,role
john@company.com,John Doe,ABC Company,Developer
jane@business.com,Jane Smith,XYZ Business,Manager
```


### Using Attributes in Emails

Insert dynamic content using double curly braces:

**Subject Example:**

```
Hello {{name}}, exciting news from {{company}}!
```

**Email Body Example:**

```
Dear {{name}},

Thank you for your interest in our services. We're excited to work with {{company}}!

Your position as {{position}} makes you an ideal candidate for our solution.

Best regards,
The Team
```


## 🔧 Configuration

### Email Provider Settings

The application auto-configures SMTP settings for major providers:


| Provider | SMTP Server | Port | Security |
| :-- | :-- | :-- | :-- |
| Gmail | smtp.gmail.com | 587 | TLS |
| Outlook | smtp-mail.outlook.com | 587 | TLS |
| Yahoo | smtp.mail.yahoo.com | 587 | TLS |
| iCloud | smtp.mail.me.com | 587 | TLS |

### Custom SMTP Configuration

For other email providers, select "Custom" and configure:

- **SMTP Server:** Your provider's SMTP server address
- **Port:** Usually 587 (TLS) or 465 (SSL)
- **Security:** Enable TLS or SSL as required
- **Username/Password:** Your email credentials


### Configuration File

Settings are automatically saved to `config.json` and include:

- SMTP server settings
- Last used CSV file path
- Email templates
- Application preferences


## 📊 Application Structure

```
EmailAutomation/
├── email_gui.py          # Main GUI application
├── email_app.py          # Backend email processing
├── config.json           # Saved configuration
├── recipients.csv        # Default CSV file
└── requirements.txt      # Python dependencies
```


## 🎯 Features Overview

### Modern User Interface

- **Sidebar Navigation:** Easy access to all functions
- **Card-Based Layout:** Clean, professional design
- **Responsive Design:** Adapts to different window sizes
- **Loading Animations:** Visual feedback for all operations
- **Progress Tracking:** Real-time campaign monitoring


### Email Composition

- **Template System:** Dynamic attribute insertion
- **Rich Text Editor:** Professional email formatting
- **Attachment Manager:** Add, remove, and preview files
- **Preview Function:** See personalized emails before sending
- **Subject Personalization:** Dynamic subject lines


### Campaign Management

- **Test Email:** Verify setup before bulk sending
- **Progress Monitoring:** Real-time sending statistics
- **Error Handling:** Comprehensive error reporting
- **Pause/Resume:** Campaign control options
- **Results Dashboard:** Success/failure analytics


### Activity Logging

- **Comprehensive Logs:** All operations are logged
- **Export Functionality:** Save logs as HTML files
- **Timestamp Tracking:** Precise operation timing
- **Error Details:** Complete error information
- **Search/Filter:** Find specific log entries


## 🔍 Troubleshooting

### Common Issues

**Connection Errors:**

- Verify email credentials are correct
- For Gmail: Use App Password, not regular password
- Check firewall/antivirus isn't blocking the application
- Ensure internet connection is stable

**CSV Loading Issues:**

- Verify CSV file format is valid
- Check file encoding (UTF-8 recommended)
- Ensure file isn't open in another application
- Try a different CSV file to isolate the issue

**Email Sending Failures:**

- Test connection before bulk sending
- Verify recipient email addresses are valid
- Check for attachment file size limits
- Monitor sending rate limits from your provider

**Application Errors:**

- Check Python version (3.7+ required)
- Verify all dependencies are installed
- Look at console output for error messages
- Try running with administrator privileges


### Getting Help

1. **Check Activity Logs:** Look for detailed error messages
2. **Test Configuration:** Use the connection test feature
3. **Verify CSV Format:** Ensure proper CSV structure
4. **Contact Support:** younes0079@gmail.com

## 📧 Contact \& Support

- **Developer:** ByeBye21
- **Email:** younes0079@gmail.com
- **GitHub:** [github.com/ByeBye21](https://github.com/ByeBye21)
