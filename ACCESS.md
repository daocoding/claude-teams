# Teams Channel Access & Delivery

The configuration file lives at `~/.claude/channels/teams/access.json`.

```jsonc
{
  "dmPolicy": "pairing",
  "allowFrom": ["19:xxxx@unq.gbl.spaces"],
  "senderMap": {
    "19:xxxx@unq.gbl.spaces": "Jane Doe"
  },
  "groups": {
    "19:yyyy@thread.v2": {
      "requireMention": false,
      "allowFrom": []
    }
  },
  "pending": {}
}
```

## Fields

| Field | Description |
|-------|-------------|
| `dmPolicy` | `"pairing"` (default), `"allowlist"`, or `"disabled"` |
| `allowFrom` | Array of approved conversation IDs |
| `senderMap` | Display names and service URLs for known conversations |
| `groups` | Group chat policies — `requireMention` gates on @mention, `allowFrom` restricts by sender |
| `pending` | Conversations awaiting approval, auto-expire after 24 hours |

## Pairing flow

1. Unknown user DMs the bot
2. Bot replies with a 6-character code (e.g., `A3X7K2`)
3. User communicates the code to the Claude Code operator out-of-band
4. Operator runs `/teams:access pair A3X7K2`
5. Conversation moves to `allowFrom`, future messages are delivered

## DM policies

- **pairing** — unknown DMs get a pairing code, known DMs are delivered
- **allowlist** — only `allowFrom` conversations are delivered, unknown DMs are dropped silently
- **disabled** — all DMs are dropped

## Group chats

Groups must be explicitly added. Each group can optionally:
- Require `@mention` to trigger delivery (`requireMention: true`)
- Restrict which senders in the group are delivered (`allowFrom` array)

All changes via `/teams:access` take effect immediately without restart.
