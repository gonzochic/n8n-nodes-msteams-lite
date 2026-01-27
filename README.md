# n8n-nodes-msteams-lite

![n8n.io - Workflow Automation](https://img.shields.io/badge/n8n-community%20node-brightgreen)
[![npm version](https://img.shields.io/npm/v/n8n-nodes-msteams-lite.svg)](https://www.npmjs.com/package/n8n-nodes-msteams-lite)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Microsoft Teams integration for n8n with **minimal permission scopes**.

## Why This Node?

Unlike full Microsoft Teams integrations that require broad access to your organization's data, this node uses only the permissions needed for your specific operations. Ideal for security-conscious organizations that need Teams automation without granting extensive access to channels, files, or organizational data.

## Why Choose This Node Over the Built-in n8n Teams Node?

### 1. Minimal Permissions / Least Privilege

- Only request the scopes you actually need
- Customize permissions based on your specific use case
- Easier approval from security teams and IT admins
- Audit-friendly permission model

### 2. Enhanced Trigger Functionality

- **"Fetch Full Message"** option automatically retrieves complete message content on trigger
- No need for a separate "Get Message" node after the trigger
- Works for both chat and channel messages

### 3. Non-Owner Role for Team Members

- Uses `TeamMember.ReadWriteNonOwnerRole.All` instead of `TeamMember.ReadWrite.All`
- Cannot accidentally promote users to team owners
- Safer for automated member management

### 4. Granular Control

- Choose exactly which features you need
- Remove unused scopes from your Azure AD app registration
- Clear mapping between actions and required permissions (see table below)

## Features

### Node Operations

| Resource | Operations |
|----------|------------|
| **Chat Message** | Create, Get, Get Many |
| **Channel** | Create |
| **Channel Message** | Create, Get, Get Many |
| **Team Member** | Add, Remove |
| **Channel Member** | Add, Remove |

### Trigger Events

| Event | Description |
|-------|-------------|
| **New Chat** | Fires when a new chat is created |
| **New Chat Message** | Watch all chats or a specific chat for new messages |
| **New Channel** | Fires when a new channel is created in a team |
| **New Channel Message** | Watch all channels or a specific channel for new messages |
| **New Team Member** | Fires when a new member joins a team |

**Note:** Chat and channel message triggers support the **"Fetch Full Message"** option to automatically retrieve complete message content.

## Permissions

### Permission-to-Action Mapping

Use this table to determine which scopes you need based on the features you want to use:

| Action | Required Scopes |
|--------|-----------------|
| **Chat Operations** | |
| Create chat message | `Chat.Create`, `ChatMessage.Send`, `User.Read.All` |
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
| Add/Remove channel member | `ChannelMember.ReadWrite.All`, `User.Read.All` |
| Trigger: New team member | `TeamMember.ReadWriteNonOwnerRole.All` |
| **Always Required** | |
| Token refresh | `offline_access` |

### Full Scope List

| Scope | Purpose |
|-------|---------|
| `Chat.Create` | Create new 1-on-1 and group chats |
| `Chat.ReadWrite` | Read and manage chat metadata, required for chat subscriptions |
| `ChatMessage.Read` | Read messages in chats |
| `ChatMessage.Send` | Send messages to chats |
| `Channel.Create` | Create channels in teams |
| `ChannelMessage.Read.All` | Read messages in channels |
| `ChannelMessage.Send` | Send messages to channels |
| `ChannelMember.ReadWrite.All` | Add and remove channel members |
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

> **Important:** This node requires its own credential type ("Reduced Permissions Microsoft Teams OAuth2 API"). You cannot use the built-in Microsoft Teams OAuth2 credentials that ship with n8n, as those require broader permission scopes.

1. **Register an Azure AD application**
   - Go to [Azure Portal](https://portal.azure.com/) > App registrations > New registration
   - Set a name and choose the appropriate account type

2. **Configure OAuth2 redirect URI**
   - In n8n, start creating the credential - the "OAuth Redirect URL" is displayed at the top
   - In Azure, add a Web platform redirect URI with this URL

3. **Add API permissions**
   - Go to API permissions > Add a permission > Microsoft Graph > Delegated permissions
   - Add only the scopes you need based on the [Permission-to-Action Mapping](#permission-to-action-mapping) table
   - For all features, add: `Chat.Create`, `Chat.ReadWrite`, `ChatMessage.Read`, `ChatMessage.Send`, `Channel.Create`, `ChannelMessage.Read.All`, `ChannelMessage.Send`, `ChannelMember.ReadWrite.All`, `TeamMember.ReadWriteNonOwnerRole.All`, `User.Read.All`, `offline_access`
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
