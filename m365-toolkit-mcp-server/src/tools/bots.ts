// ---------------------------------------------------------------------------
// Teams Bot tools
// ---------------------------------------------------------------------------

import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { ResponseFormat } from "../types.js";
import type { TeamsBotManifest } from "../types.js";
import { TEAMS_MANIFEST_SCHEMA } from "../services/knowledge.js";

// -- Helpers -----------------------------------------------------------------

function buildCommandList(
  features: string[]
): Record<string, unknown>[] {
  const commands: Record<string, unknown>[] = [];

  if (features.includes("meeting_scheduling")) {
    commands.push({
      scopes: ["personal", "team", "groupchat"],
      commands: [{ title: "Schedule Meeting", description: "Schedule a new meeting" }],
    });
  }
  if (features.includes("notes_taking")) {
    commands.push({
      scopes: ["personal", "team", "groupchat"],
      commands: [{ title: "Take Notes", description: "Start taking meeting notes" }],
    });
  }
  if (features.includes("task_management")) {
    commands.push({
      scopes: ["personal", "team", "groupchat"],
      commands: [{ title: "Create Task", description: "Create and track a task" }],
    });
  }
  if (features.includes("notifications")) {
    commands.push({
      scopes: ["personal", "team", "groupchat"],
      commands: [{ title: "Notifications", description: "Manage proactive notifications" }],
    });
  }

  if (commands.length === 0) {
    commands.push({ scopes: ["personal", "team", "groupchat"], commands: [] });
  }

  return commands;
}

// -- Schema ------------------------------------------------------------------

const CreateBotInputSchema = z.object({
  name: z.string().min(1).describe("Bot display name"),
  features: z
    .array(
      z.enum([
        "meeting_scheduling",
        "notes_taking",
        "task_management",
        "notifications",
      ])
    )
    .default([])
    .describe(
      "Bot features: meeting_scheduling, notes_taking, task_management, notifications"
    ),
  response_format: z
    .nativeEnum(ResponseFormat)
    .default(ResponseFormat.JSON)
    .describe("Output format"),
}).strict();

type CreateBotInput = z.infer<typeof CreateBotInputSchema>;

// -- Registration ------------------------------------------------------------

export function registerBotTools(server: McpServer): void {
  server.registerTool(
    "m365_create_teams_bot_config",
    {
      title: "Create Teams Bot Config",
      description: `Generate a Microsoft Teams app manifest (manifest.json) for a conversational bot.

The manifest follows the Teams schema v1.17 and includes bot configuration with command lists for the selected features.

Available features:
  - meeting_scheduling: Schedule and manage meetings
  - notes_taking: Take and share meeting notes
  - task_management: Create and track tasks
  - notifications: Send proactive notifications

Args:
  - name (string): Bot display name
  - features (string[]): Bot features to include (default: [])
  - response_format ('json' | 'markdown'): Output format (default: json)

Returns:
  Complete Teams app manifest JSON. Replace \${TEAMS_APP_ID} and \${BOT_ID} placeholders with actual values before deploying.`,
      inputSchema: CreateBotInputSchema,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params: CreateBotInput) => {
      const manifest: TeamsBotManifest = {
        $schema: TEAMS_MANIFEST_SCHEMA,
        manifestVersion: "1.17",
        version: "1.0.0",
        id: "${TEAMS_APP_ID}",
        packageName: `com.teams.${params.name.toLowerCase().replace(/\s+/g, "")}`,
        developer: {
          name: "Teams Developer",
          websiteUrl: "https://www.example.com",
          privacyUrl: "https://www.example.com/privacy",
          termsOfUseUrl: "https://www.example.com/terms",
        },
        icons: { color: "color.png", outline: "outline.png" },
        name: { short: params.name, full: params.name },
        description: {
          short: `${params.name} - A Teams bot`,
          full: `${params.name} is a bot for Microsoft Teams with advanced features`,
        },
        accentColor: "#FFFFFF",
        bots: [
          {
            botId: "${BOT_ID}",
            scopes: ["personal", "team", "groupchat"],
            supportsFiles: params.features.includes("meeting_scheduling"),
            isNotificationOnly: false,
            commandLists: buildCommandList(params.features),
          },
        ],
        permissions: ["identity", "messageTeamMembers"],
        validDomains: [],
      };

      let text: string;
      if (params.response_format === ResponseFormat.MARKDOWN) {
        const lines = [
          `# Teams Bot Manifest: ${params.name}`,
          "",
          `**Features**: ${params.features.join(", ") || "basic messaging"}`,
          "",
          "```json",
          JSON.stringify(manifest, null, 2),
          "```",
          "",
          "## Placeholders to replace",
          "- `${TEAMS_APP_ID}` → Your Teams app ID",
          "- `${BOT_ID}` → Your Azure Bot registration ID",
        ];
        text = lines.join("\n");
      } else {
        text = JSON.stringify(manifest, null, 2);
      }

      return {
        content: [{ type: "text" as const, text }],
      };
    }
  );
}
