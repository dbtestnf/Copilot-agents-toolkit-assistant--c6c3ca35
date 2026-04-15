// ---------------------------------------------------------------------------
// Agent Store template tools
// ---------------------------------------------------------------------------

import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { ResponseFormat } from "../types.js";
import {
  getStoreTemplates,
  getTemplateConversationStarters,
} from "../services/knowledge.js";

// -- Schema ------------------------------------------------------------------

const ListStoreTemplatesInputSchema = z.object({
  category: z
    .enum(["Support", "Sales", "HR", "Analytics", "Productivity"])
    .optional()
    .describe("Filter by category"),
  sort_by: z
    .enum(["downloads", "rating", "name"])
    .default("downloads")
    .describe("Sort field (default: downloads)"),
  response_format: z
    .nativeEnum(ResponseFormat)
    .default(ResponseFormat.MARKDOWN)
    .describe("Output format"),
}).strict();

type ListStoreTemplatesInput = z.infer<typeof ListStoreTemplatesInputSchema>;

// -- Registration ------------------------------------------------------------

export function registerTemplateTools(server: McpServer): void {
  server.registerTool(
    "m365_list_store_templates",
    {
      title: "List Agent Store Templates",
      description: `Browse the Microsoft 365 Agent Store catalog of pre-built agent templates.

Each template includes a name, category, description, author, rating, download count, and capabilities list.

Args:
  - category ('Support' | 'Sales' | 'HR' | 'Analytics' | 'Productivity', optional): Filter by category
  - sort_by ('downloads' | 'rating' | 'name'): Sort order (default: downloads)
  - response_format ('json' | 'markdown'): Output format (default: markdown)

Returns:
  List of available agent templates with metadata.`,
      inputSchema: ListStoreTemplatesInputSchema,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params: ListStoreTemplatesInput) => {
      let templates = Object.entries(getStoreTemplates());

      // Filter
      if (params.category) {
        templates = templates.filter(([, t]) => t.category === params.category);
      }

      // Sort
      templates.sort((a, b) => {
        if (params.sort_by === "name") return a[1].name.localeCompare(b[1].name);
        if (params.sort_by === "rating") return b[1].rating - a[1].rating;
        return b[1].downloads - a[1].downloads;
      });

      // Build output
      const output = templates.map(([id, t]) => ({
        id,
        ...t,
        conversation_starters: getTemplateConversationStarters(id),
      }));

      let text: string;
      if (params.response_format === ResponseFormat.MARKDOWN) {
        const lines = ["# Agent Store Templates", ""];
        if (params.category) lines.push(`**Category filter**: ${params.category}`, "");
        for (const item of output) {
          lines.push(`## ${item.name} (\`${item.id}\`)`);
          lines.push(
            `**Category**: ${item.category} | **Author**: ${item.author} | ⭐ ${item.rating}/5.0 | 📥 ${item.downloads.toLocaleString()} downloads`
          );
          lines.push(`${item.description}`);
          lines.push(`**Capabilities**: ${item.capabilities.join(", ")}`);
          if (item.conversation_starters.length > 0) {
            lines.push("**Conversation Starters**:");
            item.conversation_starters.forEach((s) =>
              lines.push(`- _${s.title}_: "${s.text}"`)
            );
          }
          lines.push("");
        }
        lines.push(`**Total**: ${output.length} template(s)`);
        text = lines.join("\n");
      } else {
        text = JSON.stringify(output, null, 2);
      }

      return {
        content: [{ type: "text" as const, text }],
      };
    }
  );
}
