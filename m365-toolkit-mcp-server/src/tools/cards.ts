// ---------------------------------------------------------------------------
// Adaptive Card tools
// ---------------------------------------------------------------------------

import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { ResponseFormat } from "../types.js";
import type { AdaptiveCard } from "../types.js";

// -- Card builders -----------------------------------------------------------

function buildBasicCard(title: string, description: string): AdaptiveCard {
  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      { type: "TextBlock", text: title, weight: "Bolder", size: "Large" },
      { type: "TextBlock", text: description, wrap: true },
    ],
  };
}

function buildApprovalCard(
  title: string,
  requestor: string,
  amount: string,
  description: string,
  date: string
): AdaptiveCard {
  const id = `approval-${title.toLowerCase().replace(/\s+/g, "-")}`;
  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      { type: "TextBlock", text: title, weight: "Bolder", size: "Large", color: "Accent" },
      {
        type: "FactSet",
        facts: [
          { title: "Requestor:", value: requestor },
          { title: "Amount:", value: amount },
          { title: "Date:", value: date },
        ],
      },
      { type: "TextBlock", text: description, wrap: true, spacing: "Medium" },
    ],
    actions: [
      { type: "Action.Submit", title: "Approve", style: "positive", data: { action: "approve", id } },
      { type: "Action.Submit", title: "Reject", style: "destructive", data: { action: "reject", id } },
    ],
  };
}

function buildNotificationCard(
  title: string,
  message: string,
  url?: string,
  style?: string
): AdaptiveCard {
  const card: AdaptiveCard = {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      {
        type: "Container",
        style: style ?? "emphasis",
        items: [
          { type: "TextBlock", text: title, weight: "Bolder", size: "Large" },
          { type: "TextBlock", text: message, wrap: true },
        ],
      },
    ],
  };
  if (url) {
    card.actions = [{ type: "Action.OpenUrl", title: "View Details", url }];
  }
  return card;
}

interface FormField {
  id: string;
  label: string;
  field_type: "text" | "choice";
  placeholder?: string;
  choices?: { title: string; value: string }[];
}

function buildFormCard(title: string, fields: FormField[]): AdaptiveCard {
  const body: unknown[] = [
    { type: "TextBlock", text: title, weight: "Bolder", size: "Large" },
  ];

  for (const field of fields) {
    if (field.field_type === "text") {
      body.push({
        type: "Input.Text",
        id: field.id,
        label: field.label,
        placeholder: field.placeholder ?? "",
      });
    } else if (field.field_type === "choice") {
      body.push({
        type: "Input.ChoiceSet",
        id: field.id,
        label: field.label,
        choices: field.choices ?? [],
      });
    }
  }

  return {
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    type: "AdaptiveCard",
    version: "1.5",
    body,
    actions: [{ type: "Action.Submit", title: "Submit", data: { action: "submit" } }],
  };
}

// -- Schema ------------------------------------------------------------------

const FormFieldSchema = z.object({
  id: z.string().describe("Unique field identifier"),
  label: z.string().describe("Display label"),
  field_type: z.enum(["text", "choice"]).describe("Field type"),
  placeholder: z.string().optional().describe("Placeholder text (for text fields)"),
  choices: z
    .array(z.object({ title: z.string(), value: z.string() }))
    .optional()
    .describe("Options for choice fields"),
});

const CreateCardInputSchema = z.object({
  card_type: z
    .enum(["basic", "approval", "notification", "form"])
    .describe("Type of adaptive card to create"),
  title: z.string().min(1).describe("Card title"),
  description: z.string().optional().describe("Card description / body text (for basic & approval cards)"),
  requestor: z.string().optional().describe("Requestor name (approval cards only)"),
  amount: z.string().optional().describe("Amount value (approval cards only)"),
  date: z.string().optional().describe("Date string (approval cards only, default: 'Today')"),
  message: z.string().optional().describe("Notification message (notification cards only)"),
  url: z.string().optional().describe("Details URL (notification cards only)"),
  style: z.string().optional().describe("Container style for notification cards (default: emphasis)"),
  fields: z.array(FormFieldSchema).optional().describe("Form fields (form cards only)"),
  response_format: z
    .nativeEnum(ResponseFormat)
    .default(ResponseFormat.JSON)
    .describe("Output format: 'json' or 'markdown'"),
}).strict();

type CreateCardInput = z.infer<typeof CreateCardInputSchema>;

// -- Registration ------------------------------------------------------------

export function registerCardTools(server: McpServer): void {
  server.registerTool(
    "m365_create_adaptive_card",
    {
      title: "Create Adaptive Card",
      description: `Generate a Microsoft Adaptive Card JSON (schema 1.5) for use in Teams, Outlook, and other M365 apps.

Supported card types:
  - basic: Simple text card with title and description
  - approval: Approve/Reject card with requestor, amount, date
  - notification: Alert card with message and optional URL
  - form: Data-collection card with text and choice fields

Args:
  - card_type ('basic' | 'approval' | 'notification' | 'form'): Type of card
  - title (string): Card title
  - description (string, optional): Description text
  - requestor (string, optional): For approval cards
  - amount (string, optional): For approval cards
  - date (string, optional): For approval cards (default: "Today")
  - message (string, optional): For notification cards
  - url (string, optional): For notification cards
  - style (string, optional): Container style for notification cards
  - fields (array, optional): Form fields with {id, label, field_type, placeholder?, choices?}
  - response_format ('json' | 'markdown'): Output format (default: json)

Returns:
  Complete Adaptive Card JSON ready for rendering.`,
      inputSchema: CreateCardInputSchema,
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async (params: CreateCardInput) => {
      let card: AdaptiveCard;

      switch (params.card_type) {
        case "approval":
          card = buildApprovalCard(
            params.title,
            params.requestor ?? "N/A",
            params.amount ?? "N/A",
            params.description ?? "",
            params.date ?? "Today"
          );
          break;
        case "notification":
          card = buildNotificationCard(
            params.title,
            params.message ?? "",
            params.url,
            params.style
          );
          break;
        case "form":
          card = buildFormCard(params.title, (params.fields ?? []) as FormField[]);
          break;
        default:
          card = buildBasicCard(params.title, params.description ?? "");
      }

      let text: string;
      if (params.response_format === ResponseFormat.MARKDOWN) {
        text = [
          `# Adaptive Card: ${params.title}`,
          `**Type**: ${params.card_type}`,
          "",
          "```json",
          JSON.stringify(card, null, 2),
          "```",
        ].join("\n");
      } else {
        text = JSON.stringify(card, null, 2);
      }

      return {
        content: [{ type: "text" as const, text }],
      };
    }
  );
}
