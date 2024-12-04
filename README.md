# Send-EmailFromCSV PowerShell Script

A PowerShell script that automates sending emails to users with information from a CSV file. The script features interactive approval, email address editing capabilities, secure credential handling, and customizable HTML email templates.

## Features

- Interactive email preview and approval system
- Secure SMTP credential handling with save/load functionality
- Email address editing and validation
- Support for delivery receipts and read receipts
- Customizable email priority settings
- Flexible domain configuration for email addresses
- HTML template support with variable substitution
- Dynamic subject lines from templates
- Automatic current date insertion in templates

## Prerequisites

- PowerShell Version 5.1 or later

## Parameters

| Parameter | Description | Default Value |
|-----------|-------------|---------------|
| `CsvFilePath` | Path to CSV file containing user details | `.\TempPasswords.csv` |
| `SmtpServer` | SMTP server address | `smtp.office365.com` |
| `SmtpPort` | SMTP server port | `587` |
| `SmtpUser` | Email account username for SMTP authentication | - |
| `SmtpPassword` | Email account password (SecureString) | - |
| `EmailDomain` | Domain to append to SAMAccountNames | Prompted if not provided |
| `CredentialPath` | Path for storing/loading credentials | `.\email_creds.xml` |
| `TemplateFilePath` | Path to HTML email template file | `.\email-template.html` |
| `EmailSubjectOverride` | Optional subject line to override template subject | - |
| `StoreCredential` | Switch to save credentials for future use | `false` |
| `RequestReadReceipt` | Request read receipt for emails | `false` |
| `RequestDeliveryReceipt` | Request delivery receipt for emails | `false` |
| `HighPriority` | Mark emails as high priority | `false` |

## Template Usage

The script uses HTML templates with variable substitution. Variables in the template are denoted by double curly braces: `{{variableName}}`.

Example template structure:
```html
<!DOCTYPE html>
<html>
<head>
    <title>{{subject}}</title>
</head>
<body>
    <p>Hello {{DisplayName}},</p>
    <p>Your username is: {{UserPrincipalName}}</p>
    <p>Sent on: {{SendDate}}</p>
</body>
</html>
```

Special variables:
- `{{SendDate}}` - Automatically populated with current date

## Usage Examples

### Store Credentials
```powershell
.\Send-EmailFromCSV.ps1 -StoreCredential
```

### Basic Usage with Template
```powershell
.\Send-EmailFromCSV.ps1 -TemplateFilePath ".\Templates\welcome.html"
```

### Custom Credential File
```powershell
.\Send-EmailFromCSV.ps1 -CredentialPath "C:\MyCredentials\email_creds.xml"
```

### Override Template Subject
```powershell
.\Send-EmailFromCSV.ps1 -TemplateFilePath ".\Templates\notice.html" -EmailSubjectOverride "Important Notice"
```

### Advanced Usage
```powershell
.\Send-EmailFromCSV.ps1 -CsvFilePath "C:\Users.csv" -EmailDomain "different.domain" -RequestReadReceipt -HighPriority
```

### Manual Credentials
```powershell
$cred = Get-Credential
.\Send-EmailFromCSV.ps1 -SmtpUser $cred.UserName -SmtpPassword $cred.Password
```

## Version History

- **Version 2.0** (2024-12-04)
  - Added HTML template support with variable substitution
  - Added dynamic subject line from template
  - Added current date support in templates
  - Added subject line override capability
  - Improved template processing and error handling

- **Version 1.5** (2024-12-03)
  - Added secure credential handling
  - Added email domain parameter
  - Enhanced email validation

## Author

John A. O'Neill Sr.

## Additional Resources

For more information about secure credential handling in PowerShell:
[Microsoft Documentation](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/get-credential)
