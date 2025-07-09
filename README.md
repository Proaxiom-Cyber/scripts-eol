# Exchange Online Administration Scripts

This repository is a collection of PowerShell scripts that help administrators manage and automate tasks in **Microsoft Exchange Online**. Each script is designed to be self-contained, easy to run, and well-documented so you can quickly integrate it into your workflow.

> **Current status:** This project is just getting started. `Add-MailboxAlias.ps1` is the first script to land in the repo—many more utilities will follow!

---

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Script Catalogue](#script-catalogue)
3. [Getting Started](#getting-started)
4. [Add-MailboxAlias.ps1](#add-mailboxaliasps1)
   - [Parameters](#parameters)
   - [Examples](#examples)
5. [Contributing](#contributing)
6. [License](#license)

---

## Prerequisites

The scripts in this repository are written for **Windows PowerShell 5.1** or later. You will also need:

* An account with **sufficient Exchange Online permissions** to run the desired cmdlets (e.g. `Get-Mailbox`, `Set-Mailbox`).
* The **ExchangeOnlineManagement** PowerShell module. If it is not installed, the script will attempt to install it for the current user.
* Internet connectivity to reach the Microsoft 365 endpoints.

> macOS/Linux users can run the scripts with **PowerShell 7+**; just make sure you have the Exchange Online module available and that the script is compatible with cross-platform PowerShell (contributors welcome!).

---

## Script Catalogue

| Script | Purpose |
| ------ | ------- |
| **Add-MailboxAlias.ps1** | Adds a new alias (proxy address) to a specified mailbox.

As new scripts are added, this table will grow to provide a quick at-a-glance overview.

---

## Getting Started

1. **Clone the repo** (or download the single script).

   ```powershell
   git clone https://github.com/Proaxiom-Cyber/scripts-eol.git
   cd scripts-eol
   ```

2. **Review the script** you want to run to understand what it does.
3. **Open a PowerShell session** with the necessary execution policy (e.g. `Set-ExecutionPolicy -Scope Process RemoteSigned`).
4. **Run the script** following the examples below.

> The scripts are designed to be interactive when parameters are omitted, prompting you for the required information.

---

## Add-MailboxAlias.ps1

`Add-MailboxAlias.ps1` streamlines adding a new email alias to an existing mailbox in Exchange Online. It performs environment checks, validates input, connects to Exchange Online (including support for delegated tenants), and safely applies the change.

### Parameters

| Parameter | Type | Required | Description |
| --------- | ---- | -------- | ----------- |
| `-MailboxIdentity` | `string` | No | The UPN or GUID of the target mailbox. If omitted, you will be prompted. |
| `-AliasToAdd` | `string` | No | The new alias (e.g. `sales@contoso.com`). Case-insensitive. Prompts if omitted. |
| `-EnableVerbose` | `switch` | No | Displays extra debugging information on error. |
| `-TenantId` | `string` | No | Azure AD tenant ID (GUID) to connect to when using delegated admin scenarios. |

### Examples

Add an alias interactively:

```powershell
PS C:\> ./Add-MailboxAlias.ps1
=== Exchange Online Mailbox Alias Management ===
Enter the target mailbox UPN (e.g. user@domain.com): user@contoso.com
Enter the new alias to add (e.g. alias@domain.com): sales@contoso.com
```

Specify all parameters upfront (non-interactive):

```powershell
PS C:\> ./Add-MailboxAlias.ps1 -MailboxIdentity user@contoso.com -AliasToAdd sales@contoso.com
```

Run against a delegated tenant:

```powershell
PS C:\> ./Add-MailboxAlias.ps1 -MailboxIdentity user@contoso.com -AliasToAdd sales@contoso.com -TenantId "<tenant-guid>"
```

Enable verbose error output:

```powershell
PS C:\> ./Add-MailboxAlias.ps1 -EnableVerbose
```

### What the script does

1. Verifies your PowerShell version (requires 5.1+).
2. Ensures the **ExchangeOnlineManagement** module is installed and imported.
3. Connects to Exchange Online (optionally to a specific tenant).
4. Retrieves the target mailbox and shows current email addresses.
5. Validates and adds the new alias if it does not already exist.
6. Displays the updated list of aliases.

---

## Contributing

Contributions are welcome—whether it’s new scripts, bug fixes, or documentation improvements.

1. Fork the repository and create a feature branch.
2. Adhere to the existing coding and documentation style.
3. Submit a pull request describing the change and any testing performed.

---

## License

This project is licensed under the **MIT License**. See `LICENSE` for details. 