[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', '', Justification='CredentialPath is a file path, not a credential')]
param(
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$CsvFilePath,

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$SmtpServer,

    [Parameter(Mandatory=$false)]
    [ValidateRange(1,65535)]
    [int]$SmtpPort,

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$SmtpUser,

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [System.Security.SecureString]$SmtpPassword,

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$EmailDomain,

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$CredentialPath,

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$TemplateFilePath,

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$EmailSubjectOverride,

    [Parameter(Mandatory=$false)]
    [switch]$StoreCredential,

    [Parameter(Mandatory=$false)]
    [switch]$RequestReadReceipt,

    [Parameter(Mandatory=$false)]
    [switch]$RequestDeliveryReceipt,

    [Parameter(Mandatory=$false)]
    [switch]$HighPriority
)

<#
.SYNOPSIS
    Sends emails based on CSV data using customizable HTML templates.

.DESCRIPTION
    This script reads user details from a CSV file and sends emails using a customizable
    HTML template. It displays each email in a Windows Form for approval before sending.
    The user can edit the destination email address and must approve or reject each email
    individually. Includes support for delivery receipts, read receipts, and priority settings.

    The script provides secure credential handling and supports template-based emails with
    variable substitution. Templates can be customized using HTML with {{variableName}}
    placeholders for dynamic content.

.PARAMETER CsvFilePath
    The file path of the CSV containing user details.
    Default: ".\TempPasswords.csv"

.PARAMETER SmtpServer
    The SMTP server used to send emails.
    Default: "smtp.office365.com"

.PARAMETER SmtpPort
    The port used for the SMTP server.
    Default: 587

.PARAMETER SmtpUser
    The email account username for authentication with the SMTP server.
    If not provided, will be loaded from stored credentials or prompted.

.PARAMETER SmtpPassword
    The email account password as a SecureString for authentication with the SMTP server.
    If not provided, will be loaded from stored credentials or prompted.

.PARAMETER EmailDomain
    The domain to append to SAMAccountNames when constructing email addresses.
    Default: Prompted if not provided

.PARAMETER CredentialPath
    The file path where credentials are stored/loaded.
    Default: ".\email_creds.xml"

.PARAMETER TemplateFilePath
    The path to the HTML template file for email content.
    Default: ".\email-template.html"
    Template should contain title tags for subject and use {{variableName}} for substitutions.

.PARAMETER EmailSubjectOverride
    Optional subject line that overrides the subject in the template file.

.PARAMETER StoreCredential
    Switch to save credentials for future use. When this switch is used,
    the script will prompt for credentials, save them, and then exit.

.PARAMETER RequestReadReceipt
    Switch parameter to request a read receipt for the emails.

.PARAMETER RequestDeliveryReceipt
    Switch parameter to request a delivery receipt for the emails.

.PARAMETER HighPriority
    Switch parameter to mark the emails as high priority.

.EXAMPLE
    .\Send-EmailFromCSV.ps1 -StoreCredential
    Prompts for credentials and stores them securely for future use.

.EXAMPLE
    .\Send-EmailFromCSV.ps1 -TemplateFilePath "C:\Templates\welcome.html"
    Runs the script using the specified email template file.

.EXAMPLE
    .\Send-EmailFromCSV.ps1 -TemplateFilePath "C:\Templates\notice.html" -EmailSubjectOverride "Important Notice"
    Uses a template but overrides its subject line.

.EXAMPLE
    .\Send-EmailFromCSV.ps1 -CsvFilePath "C:\Users.csv" -EmailDomain "different.domain" -RequestReadReceipt -HighPriority
    Runs the script with a custom CSV file and email domain, requesting read receipts and setting high priority.

.NOTES
    Author: John A. O'Neill Sr.
    Date: 12/01/2024
    Version: 2.0
    Change Date: 12/04/2024
    Change Purpose: Added template-based email support, allowing for customizable email content
    Prerequisite: PowerShell Version 5.1 or later

.LINK
    For more information about email templates and secure credential handling:
    https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/get-credential
#>

# Import required assemblies for Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Set default values
if (-not $CsvFilePath) { $CsvFilePath = "$PSScriptRoot\TempPasswords.csv" }
if (-not $SmtpServer) { $SmtpServer = "smtp.office365.com" }
if (-not $SmtpPort) { $SmtpPort = 587 }
if (-not $CredentialPath) { $CredentialPath = "$PSScriptRoot\email_creds.xml" }
if (-not $TemplateFilePath) { $TemplateFilePath = "$PSScriptRoot\email-template.html" }

if (-not $EmailDomain) {
    $EmailDomain = Read-Host -Prompt "Enter domain name to use when generating recipient address"
}

# Function to process email templates
function Convert-EmailTemplate {
    param (
        [string]$templateContent,
        [hashtable]$replacements
    )
    
    $processedContent = $templateContent
    foreach ($key in $replacements.Keys) {
        $processedContent = $processedContent -replace "{{$key}}", $replacements[$key]
    }
    
    return $processedContent
}

# Credential handling
if ($StoreCredential) {
    Write-Host "Storing new credentials..."
    $cred = Get-Credential -Message "Enter SMTP credentials to store"
    $cred | Export-Clixml -Path $CredentialPath
    Write-Host "Credentials stored successfully at: $CredentialPath"
    exit
}

# Handle credentials
if (-not $SmtpPassword) {
    if (Test-Path $CredentialPath) {
        Write-Host "Loading stored credentials..."
        try {
            $cred = Import-Clixml -Path $CredentialPath
            $SmtpPassword = $cred.Password
            if (-not $SmtpUser) {
                $SmtpUser = $cred.UserName
            }
            Write-Host "Credentials loaded successfully."
        }
        catch {
            Write-Error "Failed to load stored credentials: $_"
            Write-Host "Please enter credentials manually..."
            $cred = Get-Credential -Message "Enter SMTP credentials"
            $SmtpPassword = $cred.Password
            if (-not $SmtpUser) {
                $SmtpUser = $cred.UserName
            }
        }
    }
    else {
        Write-Host "No stored credentials found. Please enter credentials."
        Write-Host "Tip: Use -StoreCredential switch to save credentials for future use."
        $cred = Get-Credential -Message "Enter SMTP credentials"
        $SmtpPassword = $cred.Password
        if (-not $SmtpUser) {
            $SmtpUser = $cred.UserName
        }
    }
}

if (-not $SmtpUser) {
    Write-Error "SMTP Username is required."
    exit
}

function Show-EmailApprovalForm {
    param (
        [string]$ToEmail,
        [string]$Subject,
        [string]$Body,
        [string]$Priority,
        [bool]$HasReadReceipt=$false,
        [bool]$HasDeliveryReceipt=$false
    )

    # Create the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Email Approval"
    $form.Size = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = "CenterScreen"

    # Create labels
    $lblTo = New-Object System.Windows.Forms.Label
    $lblTo.Location = New-Object System.Drawing.Point(10, 10)
    $lblTo.Size = New-Object System.Drawing.Size(100, 20)
    $lblTo.Text = "To:"
    $form.Controls.Add($lblTo)

    # Make the email textbox editable
    $txtTo = New-Object System.Windows.Forms.TextBox
    $txtTo.Location = New-Object System.Drawing.Point(110, 10)
    $txtTo.Size = New-Object System.Drawing.Size(660, 20)
    $txtTo.ReadOnly = $false
    $txtTo.Text = $ToEmail
    $form.Controls.Add($txtTo)

    # Add email validation label
    $lblEmailValidation = New-Object System.Windows.Forms.Label
    $lblEmailValidation.Location = New-Object System.Drawing.Point(110, 32)
    $lblEmailValidation.Size = New-Object System.Drawing.Size(660, 20)
    $lblEmailValidation.ForeColor = [System.Drawing.Color]::Red
    $form.Controls.Add($lblEmailValidation)

    $lblSubject = New-Object System.Windows.Forms.Label
    $lblSubject.Location = New-Object System.Drawing.Point(10, 60)
    $lblSubject.Size = New-Object System.Drawing.Size(100, 20)
    $lblSubject.Text = "Subject:"
    $form.Controls.Add($lblSubject)

    $txtSubject = New-Object System.Windows.Forms.TextBox
    $txtSubject.Location = New-Object System.Drawing.Point(110, 60)
    $txtSubject.Size = New-Object System.Drawing.Size(660, 20)
    $txtSubject.ReadOnly = $true
    $txtSubject.Text = $Subject
    $form.Controls.Add($txtSubject)

    # Create settings group box
    $gbSettings = New-Object System.Windows.Forms.GroupBox
    $gbSettings.Location = New-Object System.Drawing.Point(10, 90)
    $gbSettings.Size = New-Object System.Drawing.Size(760, 50)
    $gbSettings.Text = "Email Settings"

    $lblPriority = New-Object System.Windows.Forms.Label
    $lblPriority.Location = New-Object System.Drawing.Point(10, 20)
    $lblPriority.Size = New-Object System.Drawing.Size(100, 20)
    $lblPriority.Text = "Priority: $Priority"
    $gbSettings.Controls.Add($lblPriority)

    $lblReceipts = New-Object System.Windows.Forms.Label
    $lblReceipts.Location = New-Object System.Drawing.Point(200, 20)
    $lblReceipts.Size = New-Object System.Drawing.Size(400, 20)
    $lblReceipts.Text = "Read Receipt: $HasReadReceipt | Delivery Receipt: $HasDeliveryReceipt"
    $gbSettings.Controls.Add($lblReceipts)

    $form.Controls.Add($gbSettings)

    # Create web browser for HTML preview
    $webBrowser = New-Object System.Windows.Forms.WebBrowser
    $webBrowser.Location = New-Object System.Drawing.Point(10, 150)
    $webBrowser.Size = New-Object System.Drawing.Size(760, 360)
    $webBrowser.DocumentText = $Body
    $form.Controls.Add($webBrowser)

    # Create buttons
    $btnApprove = New-Object System.Windows.Forms.Button
    $btnApprove.Location = New-Object System.Drawing.Point(570, 520)
    $btnApprove.Size = New-Object System.Drawing.Size(100, 30)
    $btnApprove.Text = "Approve"
    $btnApprove.Enabled = $true
    $form.Controls.Add($btnApprove)

    $btnReject = New-Object System.Windows.Forms.Button
    $btnReject.Location = New-Object System.Drawing.Point(680, 520)
    $btnReject.Size = New-Object System.Drawing.Size(100, 30)
    $btnReject.Text = "Reject"
    $form.Controls.Add($btnReject)

    # Email validation function
    $validateEmail = {
        $email = $txtTo.Text.Trim()
        $emailRegex = "^[\w-\.]+@([\w-]+\.)+[\w-]+$"  # Modified to accept any TLD
        
        if ($email -match $emailRegex) {
            $lblEmailValidation.Text = ""
            $btnApprove.Enabled = $true
            return $true
        } else {
            $lblEmailValidation.Text = "Please enter a valid email address"
            $btnApprove.Enabled = $false
            return $false
        }
    }

    # Add email validation event
    $txtTo.Add_TextChanged({
        & $validateEmail
    })

    # Initialize validation
    & $validateEmail

    # Custom DialogResult property to store the email
    $form | Add-Member -MemberType NoteProperty -Name EmailAddress -Value ""

    # Modified button click handlers
    $btnApprove.Add_Click({
        if (& $validateEmail) {
            $form.EmailAddress = $txtTo.Text.Trim()
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Yes
        }
    })

    $btnReject.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::No
    })

    # Show the form
    $result = $form.ShowDialog()

    # Return both the result and the potentially modified email address
    return @{
        Result = $result
        EmailAddress = $form.EmailAddress
    }
}

# Check if the CSV file exists and load it first
if (-Not (Test-Path -Path $CsvFilePath)) {
    Write-Error "The CSV file '$CsvFilePath' does not exist. Please provide a valid file path."
    exit
}

try {
    # Import the CSV file
    $users = Import-Csv -Path $CsvFilePath
} catch {
    Write-Error "Failed to read the CSV file. Error: $_"
    exit
}

# Now check and load the template file
if (-Not (Test-Path -Path $TemplateFilePath)) {
    Write-Error "The template file '$TemplateFilePath' does not exist. Please provide a valid template file."
    exit
}

# Read and process the template
try {
    $templateContent = Get-Content -Path $TemplateFilePath -Raw
    
    # Extract and process subject from template if not overridden
    if (-not $EmailSubjectOverride) {
        if ($templateContent -match '<title>(.*?)</title>') {
            # Process the subject line for any variables
            $subjectTemplate = $matches[1]
            $emailSubject = $subjectTemplate
            # Replace any variables in the subject using the first user's data
            if ($users.Count -gt 0) {
                $users[0].PSObject.Properties | ForEach-Object {
                    $emailSubject = $emailSubject -replace "{{$($_.Name)}}", $_.Value
                }
            }
        } else {
            $emailSubject = "New Information from IT Department"  # Default subject
        }
    } else {
        $emailSubject = $EmailSubjectOverride
    }
} catch {
    Write-Error "Failed to read or process the template file. Error: $_"
    exit
}

# Set the from email address
$fromEmail = $SmtpUser

# Initialize counters
$totalEmails = $users.Count
$sentEmails = 0
$rejectedEmails = 0

# Iterate over each user in the CSV file
foreach ($user in $users) {
    try {
        # Construct the email address using SAMAccountName and domain parameter
        $toEmail = "$($user.SamAccountName)@$EmailDomain"
        
        # Create replacements hashtable for template processing
        $replacements = @{}

        # Add all CSV properties to the replacements
        $user.PSObject.Properties | ForEach-Object {
            $replacements[$_.Name] = $_.Value
        }
        # Add current date
        $replacements['SendDate'] = (Get-Date).ToString('dddd, MMMM d, yyyy')
        
        # Process the template with the replacements
        $emailBody = Convert-EmailTemplate -templateContent $templateContent -replacements $replacements

        # Show approval form and get result
        $priority = if ($HighPriority) { "High" } else { "Normal" }
        $approvalResult = Show-EmailApprovalForm -ToEmail $toEmail -Subject $emailSubject -Body $emailBody `
            -Priority $priority -HasReadReceipt $RequestReadReceipt -HasDeliveryReceipt $RequestDeliveryReceipt

        if ($approvalResult.Result -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Create the email message using the potentially modified email address
            $message = New-Object System.Net.Mail.MailMessage
            $message.From = $fromEmail
            $message.To.Add($approvalResult.EmailAddress)
            $message.Subject = $emailSubject
            $message.Body = $emailBody
            $message.IsBodyHtml = $true

            # Set read receipt if requested
            if ($RequestReadReceipt) {
                $message.Headers.Add("Disposition-Notification-To", $fromEmail)
            }

            # Set delivery receipt if requested
            if ($RequestDeliveryReceipt) {
                $message.DeliveryNotificationOptions = [System.Net.Mail.DeliveryNotificationOptions]::OnSuccess
            }

            # Set priority if requested
            if ($HighPriority) {
                $message.Priority = [System.Net.Mail.MailPriority]::High
            }

            # Configure SMTP client
            $smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
            $smtpClient.EnableSsl = $true
            # Convert SecureString to NetworkCredential
            $smtpClient.Credentials = New-Object System.Net.NetworkCredential($SmtpUser, $SmtpPassword)

            # Send the email
            $smtpClient.Send($message)
            $sentEmails++
            Write-Host "Email approved and sent to $($approvalResult.EmailAddress)"
        }
        else {
            $rejectedEmails++
            Write-Host "Email to $toEmail was rejected by user"
        }

    } catch {
        Write-Error "Failed to process email for $toEmail. Error: $_"
    } finally {
        if ($null -ne $message) {
            $message.Dispose()
        }
    }
}

# Display summary
Write-Host "`nEmail Processing Summary:"
Write-Host "Total emails processed: $totalEmails"
Write-Host "Emails sent: $sentEmails"
Write-Host "Emails rejected: $rejectedEmails"