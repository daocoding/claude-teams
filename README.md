# Microsoft Teams Channel for Claude Code

Connect a Claude Code session to Microsoft Teams via an MCP server running on Bun. Messages from Teams are delivered to your Claude session in real time; use the `reply`, `react`, and `edit_message` tools to respond.

## Prerequisites

- [Bun](https://bun.sh) runtime: `curl -fsSL https://bun.sh/install | bash`
- An Azure Bot registration ([create one here](https://portal.azure.com/#create/Microsoft.AzureBot))
  - Copy the **App ID**, **App Password**, and **Tenant ID**
- A public HTTPS endpoint for the webhook (e.g., [ngrok](https://ngrok.com), [Tailscale Funnel](https://tailscale.com/kb/1223/funnel), or Azure App Service)

## Setup

### 1. Install the plugin

Install directly from GitHub within Claude Code:

```
/plugin install https://github.com/daocoding/claude-teams
/reload-plugins
```

Or add the MCP server manually:

```
claude mcp add teams-channel bun run start
```

### 2. Configure credentials

```
/teams:configure <APP_ID> <APP_PASSWORD> <TENANT_ID>
```

This saves credentials to `~/.claude/channels/teams/.env` (mode 0600).

### 3. Expose the webhook

The server listens on port **3980** by default. Expose it via HTTPS and register the URL in your Azure Bot's **Messaging endpoint**:

```
https://your-domain.com/api/messages
```

### 4. Launch with the channel

```
claude --channels plugin:teams-channel
```

### 5. Pair a user

1. Have the Teams user DM your bot — they'll receive a 6-character pairing code
2. In your Claude Code terminal, run: `/teams:access pair <CODE>`
3. Once all users are paired, lock it down: `/teams:access policy allowlist`

## Tools

| Tool | Description |
|------|-------------|
| `reply` | Send a message to a Teams conversation. Pass `conversation_id` from the inbound message. Optionally pass `reply_to` (activity_id) for threading. |
| `react` | React with an emoji (sent as a threaded reply — Teams Bot API has limited native reaction support). |
| `edit_message` | Update a previously sent bot message. Edits don't trigger push notifications. |

## Environment variables

| Variable | Default | Description |
|----------|---------|-------------|
| `MICROSOFT_APP_ID` | — | Azure Bot App ID (required) |
| `MICROSOFT_APP_PASSWORD` | — | Azure Bot App Password (required) |
| `MICROSOFT_TENANT_ID` | — | Azure AD Tenant ID |
| `TEAMS_WEBHOOK_PORT` | `3980` | HTTP port for Bot Framework webhook |
| `TEAMS_WEBHOOK_PATH` | `/api/messages` | Webhook URL path |
| `TEAMS_STATE_DIR` | `~/.claude/channels/teams` | State directory |

## Access control

See [ACCESS.md](ACCESS.md) for full details.

- **DM policy**: `pairing` (default) → `allowlist` (production)
- **Group chats**: must be explicitly added via `/teams:access group add <id>`
- **Pending entries**: auto-expire after 24 hours

## How it works

```
Teams User  ──Bot Framework──▶  Webhook (Bun HTTP)  ──MCP notification──▶  Claude Code
                                                                              │
Claude Code  ──MCP tool call──▶  reply/react/edit  ──Bot Framework API──▶  Teams User
```

The server runs as an MCP server connected to Claude Code via stdio. It simultaneously runs an HTTP server to receive Bot Framework webhook callbacks. Inbound messages are delivered as `notifications/claude/channel` MCP notifications; outbound messages use the Bot Framework REST API.

## Roadmap

- **Multi-user CLI sessions** — route non-owner conversations to separate `claude --print --resume` processes, so multiple Teams users can interact with the bot concurrently
- **Owner DM routing** — designate one conversation for live MCP delivery while others get independent CLI sessions
- **Group workspace isolation** — sandboxed working directories per group chat
- **File attachments** — send and receive files via Teams
- **Adaptive Cards** — rich card formatting for structured responses
- **JWT validation** — verify Bot Framework tokens for production-grade webhook security

## Limitations

- **No message history**: the Bot Framework webhook only delivers new messages. The assistant cannot retrieve earlier messages.
- **Reactions**: Teams Bot API has limited reaction support. The `react` tool sends the emoji as a threaded reply instead.
- **File attachments**: not yet supported (planned).
- **JWT validation**: the webhook does not currently validate Bot Framework JWT tokens. Use network-level security (e.g., Tailscale, Azure VNET) in the meantime.

## License

Apache-2.0
