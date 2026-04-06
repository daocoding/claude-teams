# Microsoft Teams Channel for Claude Code

Connect a Claude Code session to Microsoft Teams via an MCP server running on Bun. Messages from Teams are delivered to your Claude session in real time; use the `reply`, `react`, and `edit_message` tools to respond.

## Prerequisites

- [Bun](https://bun.sh) runtime: `curl -fsSL https://bun.sh/install | bash`
- An Azure Bot registration ([create one here](https://portal.azure.com/#create/Microsoft.AzureBot))
  - Copy the **App ID**, **App Password**, and **Tenant ID**
  - See [Bot permissions](#bot-permissions) for required scopes
- A public HTTPS endpoint for the webhook (e.g., [Tailscale Funnel](https://tailscale.com/kb/1223/funnel), VPN, or Azure App Service) — see [Security](#security-webhook-authentication)

## Setup

### 1. Clone the plugin

```bash
git clone https://github.com/daocoding/claude-teams.git
cd claude-teams
```

### 2. Configure credentials

Create the state directory and save your Azure Bot credentials:

```bash
mkdir -p ~/.claude/channels/teams
```

Then launch Claude Code and run:

```
/teams:configure <APP_ID> <APP_PASSWORD> <TENANT_ID>
```

Or write `~/.claude/channels/teams/.env` manually (see [.env.example](.env.example)).

### 3. Expose the webhook

The server listens on port **3980** by default. Expose it via HTTPS and register the URL in your Azure Bot's **Messaging endpoint**:

```
https://your-domain.com/api/messages
```

> **Important:** Do not use a raw public URL without network-level security. See [Security: Webhook Authentication](#security-webhook-authentication).

### 4. Launch with the channel

**Development / pre-approval (now):**

```bash
claude --load-development-channels /path/to/claude-teams
```

This loads the plugin directly from your local clone.

**Official plugin (future — once approved by Anthropic):**

```bash
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
| `download_attachment` | Download an attachment by URL to the local inbox. Returns the file path. |

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
- **Outbound file attachments** — send files from Claude to Teams
- **Adaptive Cards** — rich card formatting for structured responses
- **JWT validation** — verify Bot Framework tokens for production-grade webhook security

## Limitations

- **No JWT validation**: see [Security: Webhook Authentication](#security-webhook-authentication) for mitigations.
- **No message history**: the Bot Framework webhook only delivers new messages. The assistant cannot retrieve earlier messages.
- **Reactions**: Teams Bot API has limited reaction support. The `react` tool sends the emoji as a threaded reply instead.
- **File attachments**: images and files are downloaded to `~/.claude/channels/teams/inbox/`. Images can be viewed with the Read tool. Sending files from Claude to Teams is not yet supported.
## Bot Permissions

This plugin uses **minimal permissions** — only what's needed for messaging.

### Azure Bot Framework scope

The bot authenticates with `https://api.botframework.com/.default`, which grants:

- **Send and receive messages** — reply to conversations, send typing indicators
- **Edit messages** — update previously sent bot messages
- **No Graph API access** — the bot cannot read email, calendar, files, or any other Microsoft 365 data

### Teams app manifest permissions

When registering the bot in Teams Admin Center or via app manifest:

| Permission | Why |
|-----------|-----|
| `TeamMessagingSettings.Read` | Receive messages from chats |
| `ChatMessage.Send` | Send replies back |

The bot **does not** require or request:

- `User.Read.All` or any directory permissions
- `Files.Read` / `Sites.Read` or any SharePoint/OneDrive access
- `Mail.Read` / `Calendars.Read` or any Exchange access
- Admin consent (single-tenant bots only need user-level consent)

### Single-tenant vs multi-tenant

This plugin is designed for **single-tenant** deployment (your own Azure AD tenant). Set `MICROSOFT_TENANT_ID` in your `.env` to restrict authentication to your tenant only. Do not leave it empty in production — an empty tenant ID falls back to `botframework.com` which allows any tenant.

## Security: Webhook Authentication

**This plugin does not validate Bot Framework JWT tokens on incoming webhook requests.** This means anyone who discovers your webhook URL can send forged messages that will be delivered to your Claude Code session as if they came from a legitimate Teams user.

**Risk:** An attacker could impersonate an approved user or inject arbitrary prompts into your session.

**Recommended mitigations (use at least one):**

1. **Network-level security (strongly recommended)** — expose the webhook only through [Tailscale Funnel](https://tailscale.com/kb/1223/funnel), a VPN, or Azure VNET. This ensures only trusted networks can reach the endpoint.
2. **Firewall rules** — restrict inbound traffic to [Microsoft's Bot Framework IP ranges](https://learn.microsoft.com/en-us/azure/bot-service/bot-service-resources-faq-security).
3. **Reverse proxy with auth** — place the webhook behind nginx/Caddy with mutual TLS or basic auth.

**Do not expose the webhook on a public URL (e.g., raw ngrok) without one of the above.** JWT validation is on the [roadmap](#roadmap).

## License

Apache-2.0
