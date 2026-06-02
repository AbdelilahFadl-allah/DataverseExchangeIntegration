# DV_EXCHGE Integration Tool

A PowerShell-based integration tool that synchronizes email between a Dataverse / Dynamics CRM environment and Exchange Online.

## What it does

- Uses Microsoft Graph API to interact with Exchange Online.
- Uses PowerShell XRM Tooling connector and Dynamics XRM SDK libraries to interact with Dataverse.
- Synchronizes outgoing emails from Dataverse to Exchange Online.
- Synchronizes incoming emails from Exchange Online into Dataverse.
- Supports email correlation for incoming messages so that related conversations are tracked correctly.

## Configuration files

- `config.json`
  - Contains connection credentials for both Dataverse and Exchange Online.
  - Use this file to configure the service endpoints, authentication details, and any required client credentials.

- `out_mailbox.json`
  - Lists mailboxes that should be used to send email from Dataverse.
  - Each entry represents an Exchange Online mailbox configured for outgoing email synchronization.
  - Each config record includes the mailboxe's email address and the start date for synchronization in ISO format

- `in_mailbox.json`
  - Lists incoming mailboxes whose received email should be synchronized into Dataverse.
  - Each entry represents an Exchange Online mailbox to monitor for incoming messages.
  - Each config record includes the mailboxe's email address and the start date for synchronization in ISO format

## Requirements

- PowerShell environment capable of running the script.
- Azure AD application registration for Exchange Online Graph access.
- Required Graph permissions for the app registration:
  - `Mail.ReadWrite` Application Permission
  - `Mail.Read.Shared` Delegated Permission
  - `Mail.Send` Application Permission

## How to use

1. Configure `config.json` with your Dataverse and Exchange Online credentials.
2. Add outgoing mailbox settings to `out_mailbox.json`.
3. Add incoming mailbox settings to `in_mailbox.json`.
4. Run the PowerShell script .\run.ps1 to execute the synchronization.

## Notes

- The tool is designed to bridge outgoing and incoming email flows between Dynamics CRM / Dataverse and Exchange Online.
- Email correlation is supported for incoming emails so records in Dataverse can be matched to existing email threads.

## Project file tree

```text
DV_EXCHGE/
├── config.json
├── gitignore
├── in_mailbox.json
├── LoadPackages.psm1
├── logs/
├── out_mailbox.json
├── README.md
├── run.ps1
├── test_mod.psm1
└── utils.psm1
```
