# Microsoft 365 User Offboarding Script

Automate secure employee offboarding from Microsoft 365 / Entra ID.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://docs.microsoft.com/powershell/)
[![Microsoft 365](https://img.shields.io/badge/Microsoft%20365-Ready-green.svg)](https://www.microsoft.com/microsoft-365)

## What It Does

| Step | Action |
|------|--------|
| 1 | Disable sign-in |
| 2 | Reset password (random) |
| 3 | Revoke all sessions |
| 4 | Remove from all groups |
| 5 | Convert mailbox to shared |
| 6 | Set email forwarding |
| 7 | Configure auto-reply |
| 8 | Remove licenses |

## Quick Start

```powershell
# Install required modules
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser

# Run the script
.\offboarding.ps1 -UPN "user@company.com" -ManagerEmail "manager@company.com"
```

## Documentation

**📖 Full documentation is embedded in the script itself.**

Open [offboarding.ps1](offboarding.ps1) to see:
- Step-by-step explanations of each operation
- Security principles and best practices applied
- Microsoft Learn documentation links
- Customization options

## Requirements

- PowerShell 5.1+
- [Microsoft.Graph](https://www.powershellgallery.com/packages/Microsoft.Graph) module
- [ExchangeOnlineManagement](https://www.powershellgallery.com/packages/ExchangeOnlineManagement) module
- User.ReadWrite.All and Group.ReadWrite.All Graph permissions
- Exchange Administrator role (for mailbox operations)

## License

MIT License - see [LICENSE](LICENSE)

## Author

**Michael Coyle** - [@coylemichael](https://github.com/coylemichael)
