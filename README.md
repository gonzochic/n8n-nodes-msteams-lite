# n8n-nodes-msteams-lite

![n8n.io - Workflow Automation](https://img.shields.io/badge/n8n-community%20node-brightgreen)
[![npm version](https://img.shields.io/npm/v/n8n-nodes-msteams-lite.svg)](https://www.npmjs.com/package/n8n-nodes-msteams-lite)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A community node for n8n that focuses on **least-privilege permissions** and adds a few usability features on top. If your organisation requires tight control over which Microsoft Graph scopes an app registration requests, this node lets you tailor them to exactly what your workflows need.

## Why This Node?

The built-in n8n Microsoft Teams node is a great general-purpose integration that covers a wide range of Teams operations. This community node takes a different approach: it is designed for environments where **least-privilege access** is a hard requirement — for example when security teams need to review and approve every OAuth scope before an app registration goes live.

In addition to customisable scopes, it ships with a few usability improvements for common trigger-based workflows.

## How Does This Compare to the Built-in n8n Teams Node?

The built-in node covers more operations and works out of the box. This node focuses on permission control and adds trigger conveniences. Pick whichever fits your use case.

| Capability | Built-in Node | This Node |
|---|---|---|
| **OAuth2 Scopes** | Pre-defined set that covers all operations | Editable scope field — request only the scopes your workflows actually use |
| **Trigger: Fetch Full Message** | Returns notification metadata; full content can be fetched in a follow-up step | Built-in **"Fetch Full Message"** toggle retrieves complete message content in one step |
| **Trigger: Ignore Own Messages** | Not available | **"Ignore Own Messages"** toggle silently drops notifications from the authenticated user, preventing self-triggered loops |
| **Send to Existing Chat** | Not available | Chat Message Create can target an existing chat (group or one-on-one) by ID, making it easy to reply to trigger outputs |
| **Team Member Scope** | `TeamMember.ReadWrite.All` | `TeamMember.ReadWriteNonOwnerRole.All` — prevents accidental owner-role escalation |
| **Permission Transparency** | Scopes managed at the Azure AD level | Clear [permission-to-action mapping](#permission-to-action-mapping) documenting which scopes each operation requires |

## Features

### Node Operations

| Resource | Operations |
|----------|------------|
| **Chat** | Create One-on-One, Create Group |
| **Chat Message** | Create (to a user or an existing chat), Get, Get Many |
| **Channel** | Create |
| **Channel Message** | Create, Get, Get Many |
| **Team Member** | Add, Remove |

### Trigger Events

| Event | Description |
|-------|-------------|
| **New Chat** | Fires when a new chat is created |
| **New Chat Message** | Watch all chats or a specific chat for new messages |
| **New Channel** | Fires when a new channel is created in a team |
| **New Channel Message** | Watch all channels or a specific channel for new messages |
| **New Team Member** | Fires when a new member joins a team |

**Note:** Chat and channel message triggers support the **"Fetch Full Message"** option to automatically retrieve complete message content and the **"Ignore Own Messages"** option to silently drop notifications caused by the authenticated user (useful to prevent loops when a workflow both sends and listens for messages).

## Permissions

### Permission-to-Action Mapping

Use this table to determine which scopes you need based on the features you want to use:

| Action | Required Scopes |
|--------|-----------------|
| **Chat Operations** | |
| Create one-on-one / group chat | `Chat.Create`, `User.Read.All` |
| Create chat message (to a user) | `Chat.Create`, `ChatMessage.Send`, `User.Read.All` |
| Create chat message (to an existing chat) | `ChatMessage.Send` |
| Get/List chat messages | `ChatMessage.Read` |
| Trigger: New chat | `Chat.ReadWrite` |
| Trigger: New chat message | `Chat.ReadWrite` (+ `ChatMessage.Read` if using Fetch Full Message) |
| **Channel Operations** | |
| Create channel | `Channel.Create` |
| Send channel message | `ChannelMessage.Send` |
| Get/List channel messages | `ChannelMessage.Read.All` |
| Trigger: New channel | `Channel.Create` |
| Trigger: New channel message | `ChannelMessage.Read.All` |
| **Member Operations** | |
| Add/Remove team member | `TeamMember.ReadWriteNonOwnerRole.All`, `User.Read.All` |
| Trigger: New team member | `TeamMember.ReadWriteNonOwnerRole.All` |
| **Always Required** | |
| Token refresh | `offline_access` |

### Full Scope List

| Scope | Purpose |
|-------|---------|
| `Chat.Create` | Create new one-on-one and group chats |
| `Chat.ReadWrite` | Read and manage chat metadata, required for chat subscriptions |
| `ChatMessage.Read` | Read messages in chats |
| `ChatMessage.Send` | Send messages to chats |
| `Channel.Create` | Create channels in teams |
| `ChannelMessage.Read.All` | Read messages in channels |
| `ChannelMessage.Send` | Send messages to channels |
| `TeamMember.ReadWriteNonOwnerRole.All` | Add and remove team members (non-owner role only) |
| `User.Read.All` | Look up users for chat creation and member operations |
| `offline_access` | Maintain access with refresh tokens |

**Tip:** You can remove unused scopes from your Azure AD app registration based on the permission mapping table above.

## Installation

Follow the [n8n community nodes installation guide](https://docs.n8n.io/integrations/community-nodes/installation/).

```
npm install n8n-nodes-msteams-lite
```

## Credentials Setup

> **Important:** This node uses its own credential type ("Reduced Permissions Microsoft Teams OAuth2 API") with a customisable scope field. It is not compatible with the built-in Microsoft Teams OAuth2 credentials, so you will need a separate Azure AD app registration.

1. **Register an Azure AD application**
   - Go to [Azure Portal](https://portal.azure.com/) > App registrations > New registration
   - Set a name and choose the appropriate account type

2. **Configure OAuth2 redirect URI**
   - In n8n, start creating the credential - the "OAuth Redirect URL" is displayed at the top
   - In Azure, add a Web platform redirect URI with this URL

3. **Add API permissions**
   - Go to API permissions > Add a permission > Microsoft Graph > Delegated permissions
   - Add only the scopes you need based on the [Permission-to-Action Mapping](#permission-to-action-mapping) table
   - For all features, add: `Chat.Create`, `Chat.ReadWrite`, `ChatMessage.Read`, `ChatMessage.Send`, `Channel.Create`, `ChannelMessage.Read.All`, `ChannelMessage.Send`, `TeamMember.ReadWriteNonOwnerRole.All`, `User.Read.All`, `offline_access`
   - Grant admin consent if required by your organization

4. **Create client secret**
   - Go to Certificates & secrets > New client secret
   - Copy the secret value immediately (it won't be shown again)

5. **Configure in n8n**
   - Add new credentials: "Reduced Permissions Microsoft Teams OAuth2 API"
   - Enter your Client ID and Client Secret
   - Update the Authorization and Token URLs with your Tenant ID
   - Optionally adjust the Scope field to match only the permissions you configured in Azure
   - Connect your account

## Resources

- [n8n Community Nodes Documentation](https://docs.n8n.io/integrations/community-nodes/)
- [Microsoft Graph API Reference](https://learn.microsoft.com/en-us/graph/api/overview)
- [Microsoft Entra OAuth2 Documentation](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)

## License

[MIT](LICENSE.md)
