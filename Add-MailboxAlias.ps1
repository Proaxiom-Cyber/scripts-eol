#Requires -Version 5.1

param(
    [Parameter(Mandatory=$false)]
    [string]$MailboxIdentity,
    
    [Parameter(Mandatory=$false)]
    [string]$AliasToAdd,
    
    [Parameter(Mandatory=$false)]
    [switch]$EnableVerbose,

    # Optional: Specify the Azure AD tenant ID to connect to a particular tenant
    [Parameter(Mandatory=$false)]
    [string]$TenantId
)

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

# Function to validate email format
function Test-EmailFormat {
    param([string]$Email)
    try {
        $null = [System.Net.Mail.MailAddress]$Email
        return $true
    } catch {
        return $false
    }
}

# Function to check Exchange Online connection
function Test-ExchangeOnlineConnection {
    try {
        $sessions = Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"}
        return $sessions.Count -gt 0
    } catch {
        return $false
    }
}

# Main script execution
try {
    Write-ColorOutput "=== Exchange Online Mailbox Alias Management ===" "Cyan"
    
    # Check PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-ColorOutput "This script requires PowerShell 5.1 or higher. Current version: $($PSVersionTable.PSVersion)" "Red"
        exit 1
    }
    
    # Ensure Exchange Online module is installed
    Write-ColorOutput "Checking ExchangeOnlineManagement module..." "Yellow"
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-ColorOutput "Installing ExchangeOnlineManagement module..." "Cyan"
        try {
            Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser -ErrorAction Stop
            Write-ColorOutput "Module installed successfully" "Green"
        } catch {
            Write-ColorOutput "Failed to install ExchangeOnlineManagement module: $($_.Exception.Message)" "Red"
            Write-ColorOutput "Try running: Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser" "Yellow"
            exit 1
        }
    } else {
        Write-ColorOutput "ExchangeOnlineManagement module found" "Green"
    }

    # Import the module
    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Write-ColorOutput "Module imported successfully" "Green"
    } catch {
        Write-ColorOutput "Failed to import ExchangeOnlineManagement module: $($_.Exception.Message)" "Red"
        exit 1
    }

    # Check if already connected
    if (-not (Test-ExchangeOnlineConnection)) {
        if ($TenantId) {
            Write-ColorOutput "Connecting to Exchange Online (Tenant: $TenantId)..." "Cyan"
        } else {
            Write-ColorOutput "Connecting to Exchange Online..." "Cyan"
        }
        try {
            if ($TenantId) {
                # Use DelegatedOrganization to target a specific tenant (requires delegated access)
                Connect-ExchangeOnline -DelegatedOrganization $TenantId -ErrorAction Stop
            } else {
                Connect-ExchangeOnline -ErrorAction Stop
            }
            Write-ColorOutput "Connected to Exchange Online successfully" "Green"
        } catch {
            Write-ColorOutput "Failed to connect to Exchange Online: $($_.Exception.Message)" "Red"
            Write-ColorOutput "Please ensure you have proper credentials and permissions" "Yellow"
            exit 1
        }
    } else {
        Write-ColorOutput "Already connected to Exchange Online" "Green"
    }

    # Get mailbox identity if not provided
    if (-not $MailboxIdentity) {
        $MailboxIdentity = Read-Host "Enter the target mailbox UPN (e.g. user@domain.com)"
    }
    
    # Validate mailbox identity
    if ([string]::IsNullOrWhiteSpace($MailboxIdentity)) {
        Write-ColorOutput "Mailbox identity cannot be empty" "Red"
        exit 1
    }
    
    if (-not (Test-EmailFormat $MailboxIdentity)) {
        Write-ColorOutput "Warning: Mailbox identity doesn't appear to be a valid email format" "Yellow"
    }

    # Try to retrieve the mailbox
    Write-ColorOutput "Retrieving mailbox information..." "Yellow"
    try {
        $Mailbox = Get-Mailbox -Identity $MailboxIdentity -ErrorAction Stop
        Write-ColorOutput "Mailbox found: $($Mailbox.DisplayName)" "Green"
    } catch {
        Write-ColorOutput "Mailbox '${MailboxIdentity}' not found or access denied" "Red"
        Write-ColorOutput "Error: $($_.Exception.Message)" "Red"
        Write-ColorOutput "Please verify:" "Yellow"
        Write-ColorOutput "  - The mailbox exists" "Yellow"
        Write-ColorOutput "  - You have sufficient permissions" "Yellow"
        Write-ColorOutput "  - The UPN is correct" "Yellow"
        exit 1
    }

    # Show existing email addresses
    Write-ColorOutput "Current email addresses for ${MailboxIdentity}:" "Yellow"
    foreach ($email in $Mailbox.EmailAddresses) {
        Write-ColorOutput "  $email" "White"
    }

    # Get alias to add if not provided
    if (-not $AliasToAdd) {
        $AliasToAdd = Read-Host "Enter the new alias to add (e.g. alias@domain.com)"
    }
    
    # Validate alias input
    if ([string]::IsNullOrWhiteSpace($AliasToAdd)) {
        Write-ColorOutput "Alias cannot be empty" "Red"
        exit 1
    }
    
    if (-not (Test-EmailFormat $AliasToAdd)) {
        Write-ColorOutput "Warning: Alias doesn't appear to be a valid email format" "Yellow"
        $Continue = Read-Host "Do you want to continue? (y/N)"
        if ($Continue -ne "y" -and $Continue -ne "Y") {
            Write-ColorOutput "Operation cancelled by user" "Yellow"
            exit 0
        }
    }

    # Check if alias already exists (case-insensitive)
    $ExistingAliases = @()
    foreach ($email in $Mailbox.EmailAddresses) {
        if ($email -like "*$AliasToAdd" -or $email -like "*$($AliasToAdd.ToLower())" -or $email -like "*$($AliasToAdd.ToUpper())") {
            $ExistingAliases += $email
        }
    }
    
    if ($ExistingAliases.Count -gt 0) {
        Write-ColorOutput "Alias '${AliasToAdd}' already exists on this mailbox:" "Yellow"
        foreach ($alias in $ExistingAliases) {
            Write-ColorOutput "  $alias" "Yellow"
        }
    } else {
        # Add the alias
        Write-ColorOutput "Adding alias '${AliasToAdd}'..." "Cyan"
        try {
            Set-Mailbox -Identity $MailboxIdentity -EmailAddresses @{Add=$AliasToAdd} -ErrorAction Stop
            Write-ColorOutput "Alias '${AliasToAdd}' added successfully." "Green"

            # Display updated list
            $UpdatedMailbox = Get-Mailbox -Identity $MailboxIdentity
            Write-ColorOutput "Updated email addresses for ${MailboxIdentity}:" "Cyan"
            foreach ($email in $UpdatedMailbox.EmailAddresses) {
                Write-ColorOutput "  $email" "White"
            }
        } catch {
            Write-ColorOutput "Failed to add alias: $($_.Exception.Message)" "Red"
            Write-ColorOutput "Please check your permissions and try again" "Yellow"
            exit 1
        }
    }

} catch {
    Write-ColorOutput "Unexpected error occurred: $($_.Exception.Message)" "Red"
    if ($EnableVerbose) {
        Write-ColorOutput "Stack trace: $($_.ScriptStackTrace)" "Red"
    }
    exit 1
} finally {
    # Optional: Disconnect (uncomment if needed)
    # Write-ColorOutput "Disconnecting from Exchange Online..." "Yellow"
    # Disconnect-ExchangeOnline -Confirm:$false
}

Write-ColorOutput "Script completed successfully!" "Green"