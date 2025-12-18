<#
.SYNOPSIS
    Offboards a user from Microsoft 365 / Entra ID by disabling sign-in,
    revoking sessions, removing group memberships, converting the mailbox
    to shared, configuring forwarding/OOO, and removing licenses.

.DESCRIPTION
    A comprehensive PowerShell script for automating Microsoft 365 user offboarding.
    
    This script performs the following actions:
    - Disables user sign-in and resets password
    - Revokes all active sessions (immediate sign-out)
    - Removes user from all Entra ID groups
    - Converts mailbox to shared (preserving emails)
    - Sets up email forwarding to manager (optional)
    - Configures out-of-office auto-reply
    - Removes all assigned licenses

.PARAMETER UPN
    The User Principal Name (email) of the user to offboard.

.PARAMETER ManagerEmail
    The email address to forward the user's mail to. Leave blank to skip forwarding.

.EXAMPLE
    .\offboarding.ps1 -UPN "john.doe@contoso.com" -ManagerEmail "jane.manager@contoso.com"

.EXAMPLE
    .\offboarding.ps1
    # Interactive mode - prompts for UPN and manager email

.NOTES
    Author:         Michael Coyle
    GitHub:         https://github.com/coylemichael/ms-offboarding
    License:        MIT License
    Version:        1.0.0
    Last Updated:   December 2024
    
    Prerequisites:
    - Microsoft.Graph PowerShell module
    - ExchangeOnlineManagement module
    - Appropriate admin permissions in Microsoft 365

.LINK
    https://learn.microsoft.com/microsoft-365/admin/add-users/remove-former-employee

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$UPN,

    [Parameter(Mandatory = $false)]
    [string]$ManagerEmail
)

#region Error Handling Configuration
# ============================================================================
# Set strict error handling to ensure failures are caught and reported.
# Best Practice: Use terminating errors so the try/catch block can handle 
# failures properly, rather than silently continuing with partial completion.
# ============================================================================
$ErrorActionPreference = 'Stop'
#endregion

#region Module Auto-Installation
# ============================================================================
# BEST PRACTICE: Ensure required modules are available before execution
#
# This section automatically checks for and installs the required PowerShell
# modules if they are not already present, eliminating manual setup steps.
# ============================================================================
$requiredModules = @(
    @{ Name = 'Microsoft.Graph'; MinVersion = '2.0.0' }
    @{ Name = 'ExchangeOnlineManagement'; MinVersion = '3.0.0' }
)

foreach ($module in $requiredModules) {
    $installed = Get-Module -ListAvailable -Name $module.Name | 
                 Where-Object { $_.Version -ge [version]$module.MinVersion } |
                 Select-Object -First 1
    
    if (-not $installed) {
        Write-Host "Installing $($module.Name) module..." -ForegroundColor Yellow
        try {
            Install-Module -Name $module.Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Host "  $($module.Name) installed successfully." -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to install $($module.Name). Please install manually: Install-Module $($module.Name) -Scope CurrentUser"
            exit 1
        }
    }
}
#endregion

try {
    #region Interactive Input
    # Prompt interactively if parameters were not supplied
    if (-not $UPN) {
        $UPN = Read-Host "Enter the user's email (UPN) to offboard"
    }

    if (-not $ManagerEmail) {
        $ManagerEmail = Read-Host "Enter manager email for mail forwarding (leave blank to skip forwarding)"
    }
    #endregion

    #region Step 1: Connect to Microsoft Graph
    # ============================================================================
    # PRINCIPLE: Least Privilege
    # 
    # We request only the minimum scopes required for the operations performed:
    #   - User.ReadWrite.All: Update user properties, disable sign-in, reset password, remove licenses
    #   - Group.ReadWrite.All: Remove user from group memberships
    #
    # We intentionally do NOT request broader scopes like Directory.ReadWrite.All
    # or unused scopes like Mail.ReadWrite, Mail.Send.
    #
    # See: [Permissions Reference](https://learn.microsoft.com/graph/permissions-reference)
    #      [RBAC Best Practices](https://learn.microsoft.com/entra/identity/role-based-access-control/best-practices)
    # ============================================================================
    Write-Host "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "User.ReadWrite.All","Group.ReadWrite.All" -ErrorAction Stop
    #endregion

    #region Step 2: Connect to Exchange Online
    # ============================================================================
    # REQUIREMENT: Exchange Online cmdlets require a separate PowerShell session
    #
    # Cmdlets like Set-Mailbox and Set-MailboxAutoReplyConfiguration are not
    # available through Microsoft Graph - they require the Exchange Online
    # PowerShell module with an active connection.
    #
    # PRINCIPLE: Graceful Degradation
    # If Exchange connection fails (e.g., no Exchange license or permissions),
    # we continue with identity cleanup rather than failing the entire script.
    #
    # See: [Connect to Exchange Online PowerShell](https://learn.microsoft.com/powershell/exchange/connect-to-exchange-online-powershell)
    # ============================================================================
    $exchangeConnected = $false
    try {
        Write-Host "Connecting to Exchange Online..."
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $exchangeConnected = $true
    }
    catch {
        Write-Warning "Could not connect to Exchange Online. Mailbox-related steps will be skipped. Error: $($_.Exception.Message)"
        $exchangeConnected = $false
    }
    #endregion

    #region Step 3: Retrieve User Object
    # ============================================================================
    # BEST PRACTICE: Explicitly request required properties
    #
    # Microsoft Graph may not return all properties by default. By specifying
    # -Property explicitly, we ensure AssignedLicenses is available for
    # license removal later in the script.
    #
    # See: [Get-MgUser](https://learn.microsoft.com/powershell/module/microsoft.graph.users/get-mguser)
    #      [Query Parameters](https://learn.microsoft.com/graph/query-parameters#select-parameter)
    # ============================================================================
    Write-Host "Retrieving user object for $UPN ..."
    $user = Get-MgUser -UserId $UPN -Property "Id,DisplayName,UserPrincipalName,AssignedLicenses" -ErrorAction Stop
    #endregion

    #region Step 4: Check for Exchange Mailbox
    # ============================================================================
    # PRINCIPLE: Defensive Programming
    #
    # Not all users have Exchange mailboxes (e.g., service accounts, users with
    # F1 licenses, external members). Verifying mailbox existence before
    # attempting mailbox operations prevents cryptic errors.
    #
    # See: [Get-Mailbox](https://learn.microsoft.com/powershell/module/exchange/get-mailbox)
    # ============================================================================
    $mailboxExists = $false
    if ($exchangeConnected) {
        try {
            $null = Get-Mailbox -Identity $UPN -ErrorAction Stop
            $mailboxExists = $true
            Write-Host "Exchange mailbox found for $UPN."
        }
        catch {
            Write-Warning "No Exchange Online mailbox found for $UPN. Mailbox steps will be skipped."
            $mailboxExists = $false
        }
    }
    #endregion

    #region Step 5: Disable User Sign-In
    # ============================================================================
    # MICROSOFT GUIDANCE: This is Step 1 in Microsoft's offboarding documentation
    #
    # Disabling sign-in should be the FIRST action to prevent the user from
    # accessing resources or taking actions during the offboarding process.
    #
    # PRINCIPLE: Defense in Depth
    # Combined with session revocation, this ensures both new and existing
    # authentication attempts are blocked.
    #
    # See: [Remove Former Employee Step 1](https://learn.microsoft.com/microsoft-365/admin/add-users/remove-former-employee-step-1)
    #      [Revoke User Access](https://learn.microsoft.com/entra/identity/users/users-revoke-access)
    # ============================================================================
    Write-Host "Disabling user sign-in..."
    Update-MgUser -UserId $UPN -AccountEnabled:$false -ErrorAction Stop
    #endregion

    #region Step 6: Reset Password
    # ============================================================================
    # SECURITY PRINCIPLE: Never use hardcoded passwords
    #
    # Hardcoded passwords in scripts create a security vulnerability - anyone
    # with access to the script knows the password. This violates:
    #   - OWASP Top 10 (A07:2021 - Identification and Authentication Failures)
    #   - CIS Controls (Control 5 - Account Management)
    #
    # We generate a cryptographically random 20-character password with special
    # characters. The password is set and immediately discarded - never logged.
    #
    # See: [Update User](https://learn.microsoft.com/graph/api/user-update)
    #      [OWASP Secrets Management](https://cheatsheetseries.owasp.org/cheatsheets/Secrets_Management_Cheat_Sheet.html)
    # ============================================================================
    $resetPassword = $true  # Set to $false if you prefer to skip password reset during offboarding.

    if ($resetPassword) {
        Write-Host "Resetting user password with a random one-time value..."
        Add-Type -AssemblyName System.Web
        $newPassword = [System.Web.Security.Membership]::GeneratePassword(20, 4)

        $passwordProfile = @{
            ForceChangePasswordNextSignIn = $true
            Password = $newPassword
        }
        Update-MgUser -UserId $UPN -PasswordProfile $passwordProfile -ErrorAction Stop
    }
    #endregion

    #region Step 7: Revoke All Active Sessions
    # ============================================================================
    # REQUIREMENT: Immediate session termination
    #
    # Disabling sign-in only prevents NEW authentication. Existing sessions with
    # valid access/refresh tokens can remain active for up to 1 hour (default
    # token lifetime). Revoking sessions invalidates all tokens immediately.
    #
    # SDK COMPATIBILITY: We use Revoke-MgUserSignInSession, which is the current
    # documented cmdlet in Microsoft.Graph SDK. Older cmdlets like
    # Invoke-MgInvalidateUserRefreshToken may not be available in all versions.
    #
    # See: [Revoke Sign-In Sessions](https://learn.microsoft.com/graph/api/user-revokesigninsessions)
    #      [Remove Former Employee Step 1](https://learn.microsoft.com/microsoft-365/admin/add-users/remove-former-employee-step-1)
    # ============================================================================
    Write-Host "Revoking user sign-in sessions..."
    Revoke-MgUserSignInSession -UserId $UPN -ErrorAction Stop
    #endregion

    #region Step 8: Remove from All Entra ID Groups
    # ============================================================================
    # BEST PRACTICE: Complete group enumeration with -All parameter
    #
    # Microsoft Graph returns paginated results (default 100 items per page).
    # Without -All, users with many group memberships would have incomplete
    # cleanup. The -All parameter retrieves all pages automatically.
    #
    # FILTERING: We filter to @odata.type = '#microsoft.graph.group' to exclude
    # directory roles and other non-group objects from memberOf results.
    #
    # SDK PATTERN: Remove-MgGroupMemberByRef with -DirectoryObjectId is the
    # current documented pattern for removing group members.
    #
    # See: [List User MemberOf](https://learn.microsoft.com/graph/api/user-list-memberof)
    #      [Delete Group Members](https://learn.microsoft.com/graph/api/group-delete-members)
    #      [Graph Paging](https://learn.microsoft.com/graph/paging)
    # ============================================================================
    Write-Host "Removing user from Entra ID groups..."
    $groups = Get-MgUserMemberOf -UserId $UPN -All -ErrorAction Stop |
              Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' }

    foreach ($group in $groups) {
        $groupName = $group.AdditionalProperties['displayName']
        try {
            Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $user.Id -ErrorAction Stop
            Write-Host "  Removed from group: $groupName"
        }
        catch {
            # Dynamic groups, role-assignable groups, and PIM-managed groups may fail
            Write-Warning "  Failed to remove from group '$groupName': $($_.Exception.Message)"
        }
    }
    #endregion

    #region Step 9: Mailbox Operations
    # ============================================================================
    # CRITICAL: Order of Operations for Mailbox Conversion
    #
    # A user mailbox REQUIRES an Exchange Online license. If you remove the
    # license before converting to shared, the mailbox enters soft-delete
    # (30-day recovery period) and then permanent deletion.
    #
    # CORRECT ORDER:
    #   1. Convert mailbox to Shared (while license still assigned)
    #   2. Set forwarding and auto-reply
    #   3. Remove license (shared mailboxes under 50GB don't need one)
    #
    # BUSINESS CONTINUITY: Shared mailboxes preserve all emails, calendar items,
    # and contacts. Authorized users can access historical data.
    #
    # See: [Convert to Shared Mailbox](https://learn.microsoft.com/microsoft-365/admin/email/convert-user-mailbox-to-shared-mailbox)
    #      [Remove Former Employee Step 6](https://learn.microsoft.com/microsoft-365/admin/add-users/remove-former-employee-step-6)
    #      [About Shared Mailboxes](https://learn.microsoft.com/microsoft-365/admin/email/about-shared-mailboxes)
    # ============================================================================
    if ($mailboxExists) {
        # 9.1 Convert mailbox to Shared
        Write-Host "Converting mailbox to shared..."
        Set-Mailbox -Identity $UPN -Type Shared -ErrorAction Stop

        # 9.2 Forward email to manager (optional)
        if ([string]::IsNullOrWhiteSpace($ManagerEmail)) {
            Write-Host "No manager email provided. Skipping mail forwarding."
        }
        else {
            Write-Host "Setting mail forwarding to $ManagerEmail ..."
            Set-Mailbox -Identity $UPN -ForwardingSMTPAddress $ManagerEmail -DeliverToMailboxAndForward $true -ErrorAction Stop
        }

        # 9.3 Set Auto-Reply / Out-of-Office
        Write-Host "Configuring auto-reply..."
        Set-MailboxAutoReplyConfiguration -Identity $UPN -AutoReplyState Enabled `
            -InternalMessage "User has left the company." `
            -ExternalMessage "User has left the company." -ErrorAction Stop
    }
    else {
        Write-Host "Skipping mailbox conversion/forwarding/OOO because no Exchange Online mailbox was found or EXO connection failed."
    }
    #endregion

    #region Step 10: Remove Licenses
    # ============================================================================
    # SDK REQUIREMENT: Use Set-MgUserLicense cmdlet
    #
    # The correct Microsoft.Graph cmdlet for license management is Set-MgUserLicense.
    # The -RemoveLicenses parameter accepts an array of SKU IDs.
    #
    # NULL-SAFE: We check if AssignedLicenses exists and contains SKUs before
    # attempting removal. This prevents errors for users with no licenses
    # (e.g., already had licenses removed, or never licensed).
    #
    # COST OPTIMIZATION: Reclaiming licenses from offboarded users frees them
    # for assignment to new employees.
    #
    # See: [Set-MgUserLicense](https://learn.microsoft.com/powershell/module/microsoft.graph.users.actions/set-mguserlicense)
    #      [Remove Former Employee Step 7](https://learn.microsoft.com/microsoft-365/admin/add-users/remove-former-employee-step-7)
    # ============================================================================
    Write-Host "Removing assigned licenses (if any)..."
    $assignedLicenses = $user.AssignedLicenses

    if ($assignedLicenses) {
        $skuIds = $assignedLicenses.SkuId | Where-Object { $_ }

        if ($skuIds) {
            Set-MgUserLicense -UserId $UPN -RemoveLicenses $skuIds -AddLicenses @() -ErrorAction Stop
            Write-Host "  Removed license SKUs: $($skuIds -join ', ')"
        }
        else {
            Write-Host "  No license SKUs found on user object."
        }
    }
    else {
        Write-Host "  User has no assigned licenses to remove."
    }
    #endregion

    #region Completion
    Write-Host "Offboarding completed successfully for $UPN." -ForegroundColor Green
    #endregion
}
catch {
    Write-Error "Offboarding FAILED for $UPN. Review the error below and correct manually."
    Write-Error $_
}
finally {
    #region Cleanup: Disconnect Sessions
    # ============================================================================
    # BEST PRACTICE: Always disconnect from services in finally block
    #
    # This ensures connections are cleaned up even if the script fails.
    # We use -ErrorAction SilentlyContinue because disconnection errors
    # (e.g., already disconnected) are not actionable.
    # ============================================================================
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
    catch { }

    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
    catch { }
    #endregion
}
