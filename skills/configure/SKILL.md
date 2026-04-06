---
name: configure
description: Set up the Teams channel — save bot credentials and review access policy. Use when the user pastes Teams bot credentials, asks to configure Teams, or wants to check channel status.
user-invocable: true
allowed-tools:
  - Read
  - Write
  - Bash(ls *)
  - Bash(mkdir *)
---

# /teams:configure

Configure the Teams channel bot credentials.

## Usage

- `/teams:configure` — show current status
- `/teams:configure <APP_ID> <APP_PASSWORD> <TENANT_ID>` — save credentials
- `/teams:configure clear` — remove saved credentials

## What it does

Writes credentials to `~/.claude/channels/teams/.env` with mode 0600:

```
MICROSOFT_APP_ID=<APP_ID>
MICROSOFT_APP_PASSWORD=<APP_PASSWORD>
MICROSOFT_TENANT_ID=<TENANT_ID>
```

Create the directory if it doesn't exist: `mkdir -p ~/.claude/channels/teams`

## After configuring

1. Restart Claude Code or run `/reload-plugins` for changes to take effect
2. Launch with the channel flag: `claude --channels plugin:teams-channel`
3. Expose the webhook endpoint (port 3980 by default) via a public HTTPS URL
4. Register the webhook URL in your Azure Bot registration under "Messaging endpoint"

## Status display (no arguments)

When run without arguments, read `~/.claude/channels/teams/.env` and `~/.claude/channels/teams/access.json` and display:
- Whether credentials are saved (show App ID, mask password)
- Current DM policy
- Number of allowed conversations
- Number of pending pairings
- Next steps if anything is missing

## Security note

`pairing` mode is a temporary bootstrap — use it to capture conversation IDs, then switch to `allowlist` mode via `/teams:access policy allowlist`.
