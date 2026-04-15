// ---------------------------------------------------------------------------
// Knowledge-base reference tools
// ---------------------------------------------------------------------------

import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { ResponseFormat } from "../types.js";
import {
  getApiPluginTemplate,
  getInstructionsGuidelines,
  getConversationStartersGuidelines,
  getDeploymentChecklist,
  getAuthenticationTypes,
  getCommonPatterns,
} from "../services/knowledge.js";

// -- Shared schema fragment for response_format only tools -------------------

const ResponseFormatOnlySchema = z.object({
  response_format: z
    .nativeEnum(ResponseFormat)
    .default(ResponseFormat.MARKDOWN)
    .describe("Output format"),
}).strict();

type ResponseFormatOnlyInput = z.infer<typeof ResponseFormatOnlySchema>;

// -- Helper: format as text or json ------------------------------------------

function formatOutput(
  data: unknown,
  format: ResponseFormat,
  markdownRenderer: () => string
): string {
  if (format === ResponseFormat.MARKDOWN) return markdownRenderer();
  return JSON.stringify(data, null, 2);
}

// -- Registration ------------------------------------------------------------

export function registerKnowledgeTools(server: McpServer): void {
  // 1. API Plugin Template
  server.registerTool(
    "m365_get_api_plugin_template",
    {
      title: "Get API Plugin Template",
      description: `Return the official Microsoft 365 API plugin manifest template (schema v2.1).

An API plugin allows a declarative agent to call external REST APIs. This template includes the standard structure for runtimes, functions, and authentication configuration.

Args:
  - response_format ('json' | 'markdown'): Output format (default: markdown)

Returns:
  A blank API plugin manifest template.`,
      inputSchema: ResponseFormatOnlySchema,
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async (params: ResponseFormatOnlyInput) => {
      const template = getApiPluginTemplate();
      const text = formatOutput(template, params.response_format, () =>
        [
          "# API Plugin Template (v2.1)",
          "",
          "```json",
          JSON.stringify(template, null, 2),
          "```",
          "",
          "## Key Fields",
          "- **description_for_model**: Critical — this is what the AI reads to decide when to call your plugin",
          "- **runtimes[].auth.type**: `None`, `OAuthPluginVault`, or `ApiKeyPluginVault`",
          "- **runtimes[].spec.url**: URL to your OpenAPI spec",
        ].join("\n")
      );
      return { content: [{ type: "text" as const, text }] };
    }
  );

  // 2. Instructions Guidelines
  server.registerTool(
    "m365_get_instructions_guidelines",
    {
      title: "Get Instructions Guidelines",
      description: `Return best-practice guidelines for writing declarative agent instructions.

Covers recommended structure (role, capabilities, behavior, response format, action triggers, examples, constraints), max length (8000 chars), external file references, and a complete example.

Args:
  - response_format ('json' | 'markdown'): Output format (default: markdown)

Returns:
  Guidelines object with purpose, max_length, external_file, structure, and example fields.`,
      inputSchema: ResponseFormatOnlySchema,
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async (params: ResponseFormatOnlyInput) => {
      const guidelines = getInstructionsGuidelines();
      const text = formatOutput(guidelines, params.response_format, () =>
        [
          "# Agent Instructions Guidelines",
          "",
          `**Purpose**: ${guidelines.purpose}`,
          `**Max Length**: ${guidelines.max_length}`,
          `**External File**: ${guidelines.external_file}`,
          "",
          guidelines.structure,
          "",
          "## Example",
          "",
          "```",
          guidelines.example,
          "```",
        ].join("\n")
      );
      return { content: [{ type: "text" as const, text }] };
    }
  );

  // 3. Conversation Starters Guidelines
  server.registerTool(
    "m365_get_conversation_starters_guidelines",
    {
      title: "Get Conversation Starters Guidelines",
      description: `Return guidelines and examples for writing effective conversation starters for M365 declarative agents.

Includes best practices, schema format, good examples, and bad examples with explanations.

Args:
  - response_format ('json' | 'markdown'): Output format (default: markdown)

Returns:
  Guidelines with best_practices, schema, good_examples, and bad_examples.`,
      inputSchema: ResponseFormatOnlySchema,
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async (params: ResponseFormatOnlyInput) => {
      const guidelines = getConversationStartersGuidelines();
      const practices = guidelines.best_practices as string[];
      const good = guidelines.good_examples as { title: string; text: string }[];
      const bad = guidelines.bad_examples as { title: string; text: string; why_bad: string }[];

      const text = formatOutput(guidelines, params.response_format, () =>
        [
          "# Conversation Starters Guidelines",
          "",
          "## Best Practices",
          ...practices.map((p) => `- ${p}`),
          "",
          "## Good Examples",
          ...good.map((g) => `- **${g.title}**: "${g.text}"`),
          "",
          "## Bad Examples",
          ...bad.map((b) => `- **${b.title}**: "${b.text}" — _${b.why_bad}_`),
        ].join("\n")
      );
      return { content: [{ type: "text" as const, text }] };
    }
  );

  // 4. Deployment Checklist
  server.registerTool(
    "m365_get_deployment_checklist",
    {
      title: "Get Deployment Checklist",
      description: `Return the complete deployment checklist for publishing a declarative agent to Microsoft Teams.

Covers manifest validation, API plugin testing, icon preparation, package creation, Teams testing, and publishing steps.

Args:
  - response_format ('json' | 'markdown'): Output format (default: markdown)

Returns:
  Array of deployment steps, each with a step name and checklist items.`,
      inputSchema: ResponseFormatOnlySchema,
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async (params: ResponseFormatOnlyInput) => {
      const checklist = getDeploymentChecklist();
      const text = formatOutput(checklist, params.response_format, () => {
        const lines = ["# Deployment Checklist", ""];
        for (const step of checklist) {
          lines.push(`## ${step.step}`);
          step.items.forEach((item) => lines.push(`- [ ] ${item}`));
          lines.push("");
        }
        return lines.join("\n");
      });
      return { content: [{ type: "text" as const, text }] };
    }
  );

  // 5. Authentication Types
  server.registerTool(
    "m365_get_authentication_types",
    {
      title: "Get Authentication Types",
      description: `Return the supported authentication types for M365 API plugins.

Types: None (public), OAuthPluginVault (OAuth 2.0), ApiKeyPluginVault (API Key). Includes schema examples, use cases, and best-practice recommendations.

Args:
  - response_format ('json' | 'markdown'): Output format (default: markdown)

Returns:
  Dictionary of auth type name → {description, schema, use_cases, best_practice?}.`,
      inputSchema: ResponseFormatOnlySchema,
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async (params: ResponseFormatOnlyInput) => {
      const authTypes = getAuthenticationTypes();
      const text = formatOutput(authTypes, params.response_format, () => {
        const lines = ["# API Plugin Authentication Types", ""];
        for (const [name, info] of Object.entries(authTypes)) {
          lines.push(`## ${name}`);
          lines.push(info.description, "");
          lines.push("**Schema**:", "```json", JSON.stringify(info.schema, null, 2), "```");
          lines.push("", "**Use Cases**:");
          info.use_cases.forEach((u) => lines.push(`- ${u}`));
          if (info.best_practice) {
            lines.push("", `💡 **Best Practice**: ${info.best_practice}`);
          }
          lines.push("");
        }
        return lines.join("\n");
      });
      return { content: [{ type: "text" as const, text }] };
    }
  );

  // 6. Common Patterns
  server.registerTool(
    "m365_get_common_patterns",
    {
      title: "Get Common Agent Patterns",
      description: `Return common architectural patterns for M365 declarative agents.

Patterns: Data Retrieval, Task Automation, Notification, and Analysis. Each includes a description, recommended components, and example use cases.

Args:
  - response_format ('json' | 'markdown'): Output format (default: markdown)

Returns:
  Dictionary of pattern_id → {name, description, components, example_use_cases}.`,
      inputSchema: ResponseFormatOnlySchema,
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async (params: ResponseFormatOnlyInput) => {
      const patterns = getCommonPatterns();
      const text = formatOutput(patterns, params.response_format, () => {
        const lines = ["# Common Agent Patterns", ""];
        for (const [, pattern] of Object.entries(patterns)) {
          lines.push(`## ${pattern.name}`);
          lines.push(pattern.description, "");
          lines.push("**Components**:");
          pattern.components.forEach((c) => lines.push(`- ${c}`));
          lines.push("", "**Example Use Cases**:");
          pattern.example_use_cases.forEach((u) => lines.push(`- ${u}`));
          lines.push("");
        }
        return lines.join("\n");
      });
      return { content: [{ type: "text" as const, text }] };
    }
  );
}
