// ---------------------------------------------------------------------------
// Declarative-agent tools
// ---------------------------------------------------------------------------

import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { ResponseFormat } from "../types.js";
import {
  CAPABILITY_ALIAS_MAP,
  getCapabilitiesInfo,
  getDeclarativeAgentTemplate,
  validateAgentManifest,
} from "../services/knowledge.js";
import type { CapabilityConfig } from "../types.js";

// -- Shared helpers ----------------------------------------------------------

function resolveCapabilities(capabilities: string[]): CapabilityConfig[] {
  const info = getCapabilitiesInfo();
  const configs: CapabilityConfig[] = [];
  const seen = new Set<string>();

  for (const cap of capabilities) {
    const canonical = CAPABILITY_ALIAS_MAP[cap] ?? cap;
    if (seen.has(canonical)) continue;
    seen.add(canonical);
    if (canonical in info) {
      configs.push({ ...info[canonical].schema });
    }
  }
  return configs;
}

// -- Schemas -----------------------------------------------------------------

const CreateAgentInputSchema = z.object({
  name: z
    .string()
    .min(1, "Name is required")
    .max(100, "Name must not exceed 100 characters")
    .describe("Agent display name (max 100 characters)"),
  description: z
    .string()
    .min(1, "Description is required")
    .max(1000, "Description must not exceed 1000 characters")
    .describe("Agent description (max 1000 characters)"),
  capabilities: z
    .array(z.string())
    .default([])
    .describe(
      "Capabilities to enable. Accepted values: web_search, graph_connectors, onedrive_sharepoint, microsoft_graph, or canonical names like WebSearch, GraphConnectors, OneDriveAndSharePoint, MicrosoftGraph"
    ),
  instructions: z
    .string()
    .max(8000, "Instructions must not exceed 8000 characters")
    .optional()
    .describe("Custom agent instructions. If omitted, a sensible default is generated."),
  conversation_starters: z
    .array(
      z.object({
        title: z.string().max(50).describe("Short title (max 50 chars)"),
        text: z.string().describe("Prompt text sent to the agent"),
      })
    )
    .optional()
    .describe("Optional conversation starters (recommended: 3-5)"),
  response_format: z
    .nativeEnum(ResponseFormat)
    .default(ResponseFormat.JSON)
    .describe("Output format: 'json' (default) or 'markdown'"),
}).strict();

type CreateAgentInput = z.infer<typeof CreateAgentInputSchema>;

const ValidateManifestInputSchema = z.object({
  manifest: z
    .record(z.unknown())
    .describe("The declarativeAgent.json manifest object to validate"),
  response_format: z
    .nativeEnum(ResponseFormat)
    .default(ResponseFormat.JSON)
    .describe("Output format"),
}).strict();

type ValidateManifestInput = z.infer<typeof ValidateManifestInputSchema>;

const GetTemplateInputSchema = z.object({
  response_format: z
    .nativeEnum(ResponseFormat)
    .default(ResponseFormat.JSON)
    .describe("Output format"),
}).strict();

type GetTemplateInput = z.infer<typeof GetTemplateInputSchema>;

const ListCapabilitiesInputSchema = z.object({
  response_format: z
    .nativeEnum(ResponseFormat)
    .default(ResponseFormat.MARKDOWN)
    .describe("Output format"),
}).strict();

type ListCapabilitiesInput = z.infer<typeof ListCapabilitiesInputSchema>;

// -- Registration ------------------------------------------------------------

export function registerAgentTools(server: McpServer): void {
  // 1. Create declarative agent
  server.registerTool(
    "m365_create_declarative_agent",
    {
      title: "Create Declarative Agent",
      description: `Generate a complete Microsoft 365 declarative agent manifest (declarativeAgent.json) following the official schema v1.0.

Args:
  - name (string): Agent display name (max 100 chars)
  - description (string): Agent description (max 1000 chars)
  - capabilities (string[]): Capabilities to enable (e.g. "web_search", "microsoft_graph")
  - instructions (string, optional): Custom instructions (max 8000 chars)
  - conversation_starters (array, optional): Array of {title, text} objects
  - response_format ('json' | 'markdown'): Output format (default: json)

Returns:
  The complete declarativeAgent.json manifest ready for use with Teams Toolkit.`,
      inputSchema: CreateAgentInputSchema,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params: CreateAgentInput) => {
      const template = getDeclarativeAgentTemplate();
      template.name = params.name;
      template.description = params.description;
      template.capabilities = resolveCapabilities(params.capabilities);

      if (params.instructions) {
        template.instructions = params.instructions;
      } else {
        template.instructions = `You are ${params.name}, an AI agent specialized in your domain.\n\n${params.description}\n\n## Response Guidelines\n- Be helpful, accurate, and professional\n- Ask clarifying questions when needed\n- Provide specific, actionable information`;
      }

      if (params.conversation_starters) {
        template.conversation_starters = params.conversation_starters;
      }

      const output = template;

      let text: string;
      if (params.response_format === ResponseFormat.MARKDOWN) {
        const lines = [
          `# Declarative Agent: ${template.name}`,
          "",
          `**Description**: ${template.description}`,
          `**Capabilities**: ${template.capabilities.map((c) => c.name).join(", ") || "none"}`,
          "",
          "## Generated Manifest (JSON)",
          "",
          "```json",
          JSON.stringify(template, null, 2),
          "```",
        ];
        text = lines.join("\n");
      } else {
        text = JSON.stringify(output, null, 2);
      }

      return {
        content: [{ type: "text" as const, text }],
      };
    }
  );

  // 2. Validate agent manifest
  server.registerTool(
    "m365_validate_agent_manifest",
    {
      title: "Validate Agent Manifest",
      description: `Validate a Microsoft 365 declarative agent manifest against the official schema.

Checks required fields ($schema, version, name, description), field length limits, and best-practice recommendations (conversation starters, instructions).

Args:
  - manifest (object): The declarativeAgent.json object to validate
  - response_format ('json' | 'markdown'): Output format (default: json)

Returns:
  Validation result with is_valid flag, errors array, and warnings array.`,
      inputSchema: ValidateManifestInputSchema,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params: ValidateManifestInput) => {
      const result = validateAgentManifest(params.manifest);

      let text: string;
      if (params.response_format === ResponseFormat.MARKDOWN) {
        const lines = [
          result.is_valid ? "# ✅ Manifest is VALID" : "# ❌ Manifest has issues",
          "",
        ];
        if (result.errors.length > 0) {
          lines.push("## Errors");
          result.errors.forEach((e) => lines.push(`- ${e}`));
          lines.push("");
        }
        if (result.warnings.length > 0) {
          lines.push("## Warnings");
          result.warnings.forEach((w) => lines.push(`- ${w}`));
        }
        text = lines.join("\n");
      } else {
        text = JSON.stringify(result, null, 2);
      }

      return {
        content: [{ type: "text" as const, text }],
      };
    }
  );

  // 3. Get blank agent template
  server.registerTool(
    "m365_get_agent_template",
    {
      title: "Get Declarative Agent Template",
      description: `Return the official blank Microsoft 365 declarative agent template following schema v1.0.

Use this as a starting point and fill in the fields manually.

Args:
  - response_format ('json' | 'markdown'): Output format (default: json)

Returns:
  A blank declarativeAgent.json template with all standard fields.`,
      inputSchema: GetTemplateInputSchema,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params: GetTemplateInput) => {
      const template = getDeclarativeAgentTemplate();

      let text: string;
      if (params.response_format === ResponseFormat.MARKDOWN) {
        text = [
          "# Declarative Agent Template (v1.0)",
          "",
          "```json",
          JSON.stringify(template, null, 2),
          "```",
        ].join("\n");
      } else {
        text = JSON.stringify(template, null, 2);
      }

      return {
        content: [{ type: "text" as const, text }],
      };
    }
  );

  // 4. List capabilities
  server.registerTool(
    "m365_list_capabilities",
    {
      title: "List Agent Capabilities",
      description: `List all available Microsoft 365 declarative agent capabilities (WebSearch, GraphConnectors, OneDriveAndSharePoint, MicrosoftGraph) with descriptions, use cases, configuration schema, and requirements.

Args:
  - response_format ('json' | 'markdown'): Output format (default: markdown)

Returns:
  Detailed information about each capability.`,
      inputSchema: ListCapabilitiesInputSchema,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params: ListCapabilitiesInput) => {
      const caps = getCapabilitiesInfo();

      let text: string;
      if (params.response_format === ResponseFormat.MARKDOWN) {
        const lines = ["# Microsoft 365 Agent Capabilities", ""];
        for (const [name, info] of Object.entries(caps)) {
          lines.push(`## ${name}`);
          lines.push(`${info.description}`, "");
          lines.push("**Use Cases**:");
          info.use_cases.forEach((u) => lines.push(`- ${u}`));
          lines.push("", "**Configuration**:");
          lines.push("```json", JSON.stringify(info.schema, null, 2), "```");
          if (info.requirements.length > 0) {
            lines.push("", "**Requirements**:");
            info.requirements.forEach((r) => lines.push(`- ⚠️ ${r}`));
          }
          if (info.supported_scopes) {
            lines.push("", "**Supported Scopes**:");
            info.supported_scopes.forEach((s) => lines.push(`- ${s}`));
          }
          lines.push("");
        }
        text = lines.join("\n");
      } else {
        text = JSON.stringify(caps, null, 2);
      }

      return {
        content: [{ type: "text" as const, text }],
      };
    }
  );
}
