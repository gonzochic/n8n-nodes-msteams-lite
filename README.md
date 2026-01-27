# n8n-nodes-msteams-lite

![n8n.io - Workflow Automation](https://img.shields.io/badge/n8n-community%20node-brightgreen)
[![npm version](https://img.shields.io/npm/v/n8n-nodes-msteams-lite.svg)](https://www.npmjs.com/package/n8n-nodes-msteams-lite)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Microsoft Teams integration for n8n with **minimal permission scopes**.

## Why This Node?

Unlike full Microsoft Teams integrations that require broad access to your organization's data, this node uses only the permissions needed for chat operations. Ideal for security-conscious organizations that need Teams chat automation without granting extensive access to channels, files, or organizational data.

## Features

### Node Operations
- **Create chat message** - Send messages to 1-on-1 or group chats (auto-creates chat if needed)
- **Get message** - Retrieve a specific message from a chat
- **Get many messages** - List messages with pagination support
- **Send and wait for response** - Send a message and wait for a reply

### Trigger Events
- **New chat created** - Fires when a new chat is created
- **New chat message** - Watch all chats or a specific chat for new messages

## Permissions

This node requires only 6 Microsoft Graph scopes:

| Scope | Purpose |
|-------|---------|
| `Chat.Create` | Create new 1-on-1 and group chats |
| `Chat.ReadWrite` | Read and manage chat metadata |
| `ChatMessage.Read` | Read messages in chats |
| `ChatMessage.Send` | Send messages to chats |
| `User.Read.All` | Look up users for chat creation |
| `offline_access` | Maintain access with refresh tokens |

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
   - Add: `Chat.Create`, `Chat.ReadWrite`, `ChatMessage.Read`, `ChatMessage.Send`, `User.Read.All`, `offline_access`
   - Grant admin consent if required by your organization

4. **Create client secret**
   - Go to Certificates & secrets > New client secret
   - Copy the secret value immediately (it won't be shown again)

5. **Configure in n8n**
   - Add new credentials: "Reduced Permissions Microsoft Teams OAuth2 API"
   - Enter your Client ID and Client Secret
   - Update the Authorization and Token URLs with your Tenant ID
   - Connect your account

## Resources

- [n8n Community Nodes Documentation](https://docs.n8n.io/integrations/community-nodes/)
- [Microsoft Graph API Reference](https://learn.microsoft.com/en-us/graph/api/overview)
- [Microsoft Entra OAuth2 Documentation](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)

## License

[MIT](LICENSE.md)
