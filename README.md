# GroupClone

Copies **group memberships**, **group ownerships**, and **Exchange Online mailbox delegations** (FullAccess, SendAs, SendOnBehalf) from one user to another using Microsoft Graph and Exchange Online PowerShell.

Built for onboarding, role transitions, and account migrations where a new user needs the same access as an existing one.

## Features

- **Interactive or CLI** — prompts for source/target UPN if not passed as parameters
- **`-WhatIf` mode** — simulates all changes without applying them
- **CSV export** — every action (copied, failed, WhatIf) is logged to a timestamped CSV
- **Minimal Graph modules** — imports only `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`, and `Microsoft.Graph.Groups`
- **Device code auth for EXO** — avoids WAM broker issues on PowerShell 7.x
- **PS 7.6+ compatible**

## Prerequisites

| Requirement | Details |
|---|---|
| PowerShell | 7.6 or later |
| Modules | `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`, `Microsoft.Graph.Groups`, `ExchangeOnlineManagement` (auto-installed if missing) |
| Graph permissions | `User.Read.All`, `Group.ReadWrite.All`, `Directory.ReadWrite.All` |
| EXO role | Sufficient Exchange Online admin rights for `Get-Mailbox`, `Get-MailboxPermission`, `Add-MailboxPermission`, `Get-RecipientPermission`, `Add-RecipientPermission`, `Set-Mailbox` |

## Usage

### Interactive
```powershell
.\GroupClone.ps1
```
You'll be prompted for the source and target UPN.

### With parameters
```powershell
.\GroupClone.ps1 -SourceUser source@contoso.com -TargetUser target@contoso.com
```

### Dry run
```powershell
.\GroupClone.ps1 -SourceUser source@contoso.com -TargetUser target@contoso.com -WhatIf
```

## CSV Output

Every run produces `GroupClone_Results_<timestamp>.csv` in the working directory with these columns:

| Column | Description |
|---|---|
| `GroupId` | Object ID or mailbox identity |
| `GroupName` | Display name |
| `Type` | `Membership`, `Ownership`, `FullAccess`, `SendAs`, or `SendOnBehalf` |
| `Status` | `Copied`, `Failed`, or `WhatIf` |
| `Error` | Error message (blank on success / WhatIf) |

## Contributing

See [CONTRIBUTING.md](./CONTRIBUTING.md).

## Code of Conduct

See [CODE_OF_CONDUCT.md](./CODE_OF_CONDUCT.md).

## Security

See [SECURITY.md](./SECURITY.md).

## License

[MIT](./LICENSE)
