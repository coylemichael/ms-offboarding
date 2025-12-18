# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2025-12-18

### Added
- Initial release of Microsoft 365 User Offboarding Script
- Automatic module installation (Microsoft.Graph, ExchangeOnlineManagement)
- Disable user sign-in with immediate effect
- Secure random password reset (no hardcoded passwords)
- Revoke all active sessions across devices
- Remove user from all Entra ID groups (with pagination support)
- Convert user mailbox to shared mailbox (preserves data)
- Optional email forwarding to manager
- Automatic out-of-office reply configuration
- License removal and reclamation
- Comprehensive inline documentation with Microsoft Learn links
- Graceful error handling and degradation
- MIT License

### Security
- Follows least privilege principle for Graph API scopes
- Implements defense in depth (disable + revoke + password reset)
- Correct order of operations (convert mailbox before removing license)
- No secrets or credentials stored in script

[1.0.0]: https://github.com/coylemichael/ms-offboarding/releases/tag/v1.0.0
