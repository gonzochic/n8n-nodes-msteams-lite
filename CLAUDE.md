# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

n8n community node package providing Microsoft Teams integration with minimal permission scopes. Built with TypeScript, compiled to CommonJS for n8n runtime.

## Build & Development Commands

```bash
npm run dev          # Start n8n with node loaded + hot reload (http://localhost:5678)
npm run build        # Compile TypeScript to dist/
npm run build:watch  # Build in watch mode
npm run lint         # Check code with ESLint
npm run lint:fix     # Auto-fix linting issues
npm run release      # Create new release via release-it
```

## Architecture

### Entry Points
- `nodes/MicrosoftTeamsLite/MicrosoftTeamsLite.node.ts` - Main node for chat operations
- `nodes/MicrosoftTeamsLite/MicrosoftTeamsLiteTrigger.node.ts` - Webhook trigger node
- `credentials/MicrosoftTeamsLiteOAuth2Api.credentials.ts` - OAuth2 with minimal scopes

### Node Structure (n8n pattern)
```
nodes/MicrosoftTeamsLite/
├── actions/
│   ├── router.ts              # Routes requests to operations
│   ├── chatMessage/           # ChatMessage resource operations
│   │   ├── create.operation.ts
│   │   ├── get.operation.ts
│   │   └── getAll.operation.ts
│   └── versionDescription.ts  # Node UI configuration
├── descriptions/              # UI field definitions
├── helpers/
│   ├── types.ts               # TypeScript interfaces
│   ├── utils.ts               # Message/list utilities
│   └── utils-trigger.ts       # Webhook subscription management
├── methods/
│   └── listSearch.ts          # Dynamic dropdown loaders (ILoadOptionsFunctions)
└── transport/
    └── index.ts               # Microsoft Graph API requests with retry logic
```

### Key Patterns
- Operations follow pattern: `{resource}.operation.ts` with `execute` function
- UI properties defined in `descriptions/` and `actions/{resource}/index.ts`
- Transport layer handles all Graph API calls via `microsoftApiRequest`
- Load methods populate dynamic dropdowns (chats, teams, channels, etc.)

## Microsoft Graph API

- Base URL: `https://graph.microsoft.com/v1.0/`
- OAuth2 scopes: `Chat.Create`, `ChatMessage.Read`, `ChatMessage.Send`, `Chat.ReadWrite`, `User.Read.All`, `offline_access`
- Webhooks for trigger node use subscription model

## Testing

No automated tests configured. Test manually:
1. Run `npm run dev`
2. Create test workflows in n8n UI
3. Execute and verify behavior

## Code Style

- ESLint: `@n8n/node-cli/eslint` preset
- Prettier: tabs, single quotes, semicolons, trailing commas, 100 char width
