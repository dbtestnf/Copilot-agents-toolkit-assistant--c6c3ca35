#!/usr/bin/env node
/**
 * MCP Server for Microsoft 365 Agent Toolkit.
 *
 * Provides 14 tools for generating declarative agents, adaptive cards,
 * Teams bot manifests, workflows, and accessing the M365 knowledge base.
 *
 * Transport: stdio (local development tool)
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { SERVER_NAME, SERVER_VERSION } from "./constants.js";

// Tool registration modules
import { registerAgentTools } from "./tools/agents.js";
import { registerCardTools } from "./tools/cards.js";
import { registerWorkflowTools } from "./tools/workflows.js";
import { registerBotTools } from "./tools/bots.js";
import { registerTemplateTools } from "./tools/templates.js";
import { registerKnowledgeTools } from "./tools/knowledge.js";

// ---------------------------------------------------------------------------
// Server setup
// ---------------------------------------------------------------------------

const server = new McpServer({
  name: SERVER_NAME,
  version: SERVER_VERSION,
});

// Register all tools
registerAgentTools(server);
registerCardTools(server);
registerWorkflowTools(server);
registerBotTools(server);
registerTemplateTools(server);
registerKnowledgeTools(server);

// ---------------------------------------------------------------------------
// Start
// ---------------------------------------------------------------------------

async function main(): Promise<void> {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error(`${SERVER_NAME} v${SERVER_VERSION} running via stdio`);
}

main().catch((error: unknown) => {
  console.error("Server error:", error);
  process.exit(1);
});
