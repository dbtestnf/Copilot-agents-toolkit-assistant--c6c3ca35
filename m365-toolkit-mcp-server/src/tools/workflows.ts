// ---------------------------------------------------------------------------
// Workflow tools
// ---------------------------------------------------------------------------

import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { ResponseFormat } from "../types.js";
import type { WorkflowConfig } from "../types.js";
import { TRIGGER_MAP, ACTION_MAP } from "../services/knowledge.js";

// -- Schema ------------------------------------------------------------------

const CreateWorkflowInputSchema = z.object({
  name: z.string().min(1).describe("Workflow name"),
  triggers: z
    .array(
      z.enum([
        "document_uploaded",
        "email_received",
        "schedule",
        "form_submitted",
      ])
    )
    .min(1, "At least one trigger is required")
    .describe(
      "Trigger types: document_uploaded, email_received, schedule, form_submitted"
    ),
  actions: z
    .array(
      z.enum([
        "send_notification",
        "request_approval",
        "create_task",
        "update_database",
        "send_teams_message",
      ])
    )
    .min(1, "At least one action is required")
    .describe(
      "Action types: send_notification, request_approval, create_task, update_database, send_teams_message"
    ),
  response_format: z
    .nativeEnum(ResponseFormat)
    .default(ResponseFormat.JSON)
    .describe("Output format"),
}).strict();

type CreateWorkflowInput = z.infer<typeof CreateWorkflowInputSchema>;

// -- Registration ------------------------------------------------------------

export function registerWorkflowTools(server: McpServer): void {
  server.registerTool(
    "m365_create_workflow",
    {
      title: "Create Workflow",
      description: `Generate a Power Automate-compatible workflow configuration with triggers and actions.

Available triggers:
  - document_uploaded: When a file is uploaded to SharePoint
  - email_received: When an email arrives in Inbox
  - schedule: On a recurring schedule (daily)
  - form_submitted: When a form is submitted

Available actions:
  - send_notification: Send an email notification
  - request_approval: Start an approval flow
  - create_task: Create a task in task list
  - update_database: Update a database row
  - send_teams_message: Post to a Teams channel

Args:
  - name (string): Workflow name
  - triggers (string[]): One or more trigger types
  - actions (string[]): One or more action types
  - response_format ('json' | 'markdown'): Output format (default: json)

Returns:
  Complete workflow configuration JSON.`,
      inputSchema: CreateWorkflowInputSchema,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params: CreateWorkflowInput) => {
      const triggers = params.triggers
        .filter((t) => t in TRIGGER_MAP)
        .map((t) => ({ ...TRIGGER_MAP[t] }));

      const actions = params.actions
        .filter((a) => a in ACTION_MAP)
        .map((a) => ({ ...ACTION_MAP[a] }));

      const workflow: WorkflowConfig = {
        name: params.name,
        description: `Automated workflow: ${params.name}`,
        triggers,
        actions,
        enabled: true,
      };

      let text: string;
      if (params.response_format === ResponseFormat.MARKDOWN) {
        const lines = [
          `# Workflow: ${params.name}`,
          "",
          `**Triggers**: ${params.triggers.join(", ")}`,
          `**Actions**: ${params.actions.join(", ")}`,
          `**Enabled**: true`,
          "",
          "## Configuration",
          "",
          "```json",
          JSON.stringify(workflow, null, 2),
          "```",
          "",
          "## Next Steps",
          "1. Review and customize the configuration",
          "2. Import into Power Automate",
          "3. Configure connections and credentials",
          "4. Test the workflow",
          "5. Enable in production",
        ];
        text = lines.join("\n");
      } else {
        text = JSON.stringify(workflow, null, 2);
      }

      return {
        content: [{ type: "text" as const, text }],
      };
    }
  );
}
