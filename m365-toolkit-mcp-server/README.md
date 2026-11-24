# m365-toolkit-mcp-server

An MCP (Model Context Protocol) server that exposes the Microsoft 365 Agent Toolkit Assistant as a set of tools any LLM client can use.

## Features

**14 tools** organised into six domains:

### Declarative Agents
| Tool | Description |
|------|-------------|
| `m365_create_declarative_agent` | Generate a complete `declarativeAgent.json` manifest |
| `m365_validate_agent_manifest` | Validate a manifest against the official schema |
| `m365_get_agent_template` | Return the blank v1.0 template |
| `m365_list_capabilities` | List available capabilities (WebSearch, Graph, etc.) |

### Adaptive Cards
| Tool | Description |
|------|-------------|
| `m365_create_adaptive_card` | Create basic, approval, notification, or form cards |

### Workflows
| Tool | Description |
|------|-------------|
| `m365_create_workflow` | Generate a Power Automate-compatible workflow config |

### Teams Bots
| Tool | Description |
|------|-------------|
| `m365_create_teams_bot_config` | Generate a Teams app manifest for a bot |

### Template Store
| Tool | Description |
|------|-------------|
| `m365_list_store_templates` | Browse the Agent Store catalog |

### Knowledge Base
| Tool | Description |
|------|-------------|
| `m365_get_api_plugin_template` | Blank API plugin manifest |
| `m365_get_instructions_guidelines` | Best practices for agent instructions |
| `m365_get_conversation_starters_guidelines` | Guidelines for conversation starters |
| `m365_get_deployment_checklist` | Deployment checklist |
| `m365_get_authentication_types` | Supported auth types for API plugins |
| `m365_get_common_patterns` | Common agent patterns |

## Quick Start

```bash
# Install dependencies
npm install

# Build
npm run build

# Run via stdio (used by MCP clients)
npm start
```

## Usage with MCP Clients

### Warp

Add to your Warp MCP configuration:

```json
{
  "mcpServers": {
    "m365-toolkit": {
      "command": "node",
      "args": ["<path-to>/m365-toolkit-mcp-server/dist/index.js"]
    }
  }
}
```

### Claude Desktop

Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "m365-toolkit": {
      "command": "node",
      "args": ["<path-to>/m365-toolkit-mcp-server/dist/index.js"]
    }
  }
}
```

## Development

```bash
# Watch mode with auto-reload
npm run dev

# Clean build
npm run clean && npm run build
```

## Architecture

```
src/
├── index.ts              # Entry point & stdio transport
├── constants.ts          # Shared constants
├── types.ts              # TypeScript interfaces & enums
├── tools/
│   ├── agents.ts         # Declarative agent tools
│   ├── cards.ts          # Adaptive card tools
│   ├── workflows.ts      # Workflow tools
│   ├── bots.ts           # Teams bot tools
│   ├── templates.ts      # Template store tools
│   └── knowledge.ts      # Knowledge-base reference tools
└── services/
    └── knowledge.ts      # M365 knowledge base (data layer)
```

## License

MIT — see the repository root LICENSE file.
