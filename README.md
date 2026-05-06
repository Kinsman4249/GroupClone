# GroupClone

**Version:** 1.0.0

GroupClone copies **Microsoft Entra ID group memberships**, **group ownerships**, and **Exchange Online mailbox delegations** from one user to another using a **Graph + Exchange Online hybrid approach**.

It is designed for onboarding, role transitions, and account migrations where a new user needs the same access as an existing one.

---

## Features

- **Hybrid Graph + EXO handling**
  - Security groups and Microsoft 365 (Unified) groups are processed via Microsoft Graph
  - Mail-enabled distribution lists and mail-enabled security groups are automatically deferred and handled via Exchange Online
- **Mailbox delegation cloning**
  - FullAccess
  - SendAs
  - SendOnBehalf
- **Interactive or parameter-driven**
  - Prompts for source/target UPNs if not supplied
- **`-WhatIf` support**
  - Safely simulate all changes without applying them
- **`-ShowDebug` flag**
  - Enables detailed enumeration and decision-making output (Graph vs EXO, skipped objects, resolved group types)
- **CSV reporting**
  - Every action is logged with status and error details
- **Minimal Graph module usage**
  - Imports only:
    - `Microsoft.Graph.Authentication`
    - `Microsoft.Graph.Users`
    - `Microsoft.Graph.Groups`
- **PowerShell 7.6+ compatible**
- **EXO device-code authentication**
  - Avoids WAM / broker issues on PowerShell 7.x

---

## Prerequisites

| Requirement | Details |
|------------|---------|
| PowerShell | 7.6 or later |
| Modules | Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups, ExchangeOnlineManagement (auto-installed if missing) |
| Graph permissions | `User.Read.All`, `Group.ReadWrite.All`, `Directory.ReadWrite.All` |
| Exchange Online role | Sufficient rights to run `Get-Mailbox`, `Get-MailboxPermission`, `Add-MailboxPermission`, `Get-RecipientPermission`, `Add-RecipientPermission`, `Set-Mailbox`, `Add-DistributionGroupMember`, `Set-DistributionGroup` |

---

## Usage

### Interactive

```powershell
.\GroupClone.ps1
```

You will be prompted for the source and target user UPNs.

### With parameters

```powershell
.\GroupClone.ps1 -SourceUser source@contoso.com -TargetUser target@contoso.com
```

### Dry run (recommended)

```powershell
.\GroupClone.ps1 -SourceUser source@contoso.com -TargetUser target@contoso.com -WhatIf -ShowDebug
```

This will:
- Enumerate all groups and mailbox permissions
- Show exactly which objects are handled by Graph vs EXO
- Make **no changes**

---

## How Group Handling Works

| Object Type | Handling Method |
|------------|-----------------|
| Security groups | Microsoft Graph |
| Microsoft 365 (Unified) groups | Microsoft Graph |
| Dynamic groups | Enumerated (membership changes will fail by design) |
| Mail-enabled distribution lists | Exchange Online |
| Mail-enabled security groups | Exchange Online |
| Group ownership | Graph or EXO depending on group type |

---

## Mailbox Delegations Copied

For every mailbox where the **source user** has permissions, the same permissions are applied to the **target user**:

- FullAccess
- SendAs
- SendOnBehalf

Applicable to:
- User mailboxes
- Shared mailboxes
- Room mailboxes
- Equipment mailboxes

---

## CSV Output

Each run generates a timestamped report:

```
GroupClone_Results_YYYYMMDD_HHMMSS.csv
```

### Columns

| Column | Description |
|------|------------|
| Name | Group or mailbox display name |
| Action | Graph Member, Graph Owner, EXO Membership, EXO Ownership, FullAccess, SendAs, SendOnBehalf |
| Status | Copied, Failed, or WhatIf |
| Error | Error message (blank on success / WhatIf) |

---

## Contributing

See `CONTRIBUTING.md`.

---

## Code of Conduct

See `CODE_OF_CONDUCT.md`.

---

## Security

See `SECURITY.md`.

---

## License

MIT
