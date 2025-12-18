<div align="center">

# 🔐 Microsoft 365 User Offboarding

**Automate secure employee offboarding from Microsoft 365 / Entra ID**

[![MIT License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)](https://docs.microsoft.com/powershell/)
[![Microsoft 365](https://img.shields.io/badge/Microsoft_365-Ready-0078D4.svg)](https://www.microsoft.com/microsoft-365)

</div>

## ⚡ Quick Start

```powershell
.\offboarding.ps1 -UPN "user@company.com"
```

**That's it.** Modules install automatically. Add `-ManagerEmail "manager@company.com"` to enable email forwarding.

## 🎯 What It Does

```
✓ Disable sign-in           ✓ Remove from all groups
✓ Reset password (random)   ✓ Convert mailbox → shared
✓ Revoke all sessions       ✓ Set auto-reply & forwarding
✓ Remove licenses
```

## 📖 Documentation

**Everything is documented in the script itself** — open [`offboarding.ps1`](offboarding.ps1) to see:

- Step-by-step explanations with security principles
- Microsoft Learn links for each operation
- Customization options

## 📋 Requirements

| Requirement | Details |
|-------------|---------|
| PowerShell | 5.1+ |
| Graph Permissions | `User.ReadWrite.All`, `Group.ReadWrite.All` |
| Exchange Role | Exchange Administrator |

> 💡 Modules (`Microsoft.Graph`, `ExchangeOnlineManagement`) install automatically on first run.

<div align="center">

**[View Script](offboarding.ps1)** · **[MIT License](LICENSE)**

</div>
