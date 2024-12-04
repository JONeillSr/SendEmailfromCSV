# Send-EmailFromCSV PowerShell Script

A PowerShell script that automates sending emails to users with information from a CSV file. The script features interactive approval, email address editing capabilities, and secure credential handling.

## Features

- Interactive email preview and approval system
- Secure SMTP credential handling with save/load functionality
- Email address editing and validation
- Support for delivery receipts and read receipts
- Customizable email priority settings
- Flexible domain configuration for email addresses

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
| `EmailDomain` | Domain to append to SAMAccountNames | `nationalmanufacturing.group` |
| `CredentialPath` | Path for storing/loading credentials | `.\email_creds.xml` |
| `StoreCredential` | Switch to save credentials for future use | `false` |
| `RequestReadReceipt` | Request read receipt for emails | `false` |
| `RequestDeliveryReceipt` | Request delivery receipt for emails | `false` |
| `HighPriority` | Mark emails as high priority | `false` |

## Usage Examples

### Store Credentials
```powershell
.\Send-EmailFromCSV.ps1 -StoreCredential
```

### Basic Usage
```powershell
.\Send-EmailFromCSV.ps1
```

### Custom Credential File
```powershell
.\Send-EmailFromCSV.ps1 -CredentialPath "C:\MyCredentials\email_creds.xml"
```

### Override Stored Username
```powershell
.\Send-EmailFromCSV.ps1 -SmtpUser "different.user@domain.com"
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

- **Version 1.5** (2024-12-03)
  - Added secure credential handling
  - Added email domain parameter
  - Enhanced email validation

## Author

John A. O'Neill Sr.

## Additional Resources

For more information about secure credential handling in PowerShell:
[Microsoft Documentation](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/get-credential)
