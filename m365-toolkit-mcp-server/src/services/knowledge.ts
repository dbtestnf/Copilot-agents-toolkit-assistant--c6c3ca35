// ---------------------------------------------------------------------------
// Ported from src/utils/knowledge.py — authoritative M365 Agents Toolkit data
// ---------------------------------------------------------------------------

import type {
  AgentPattern,
  ApiPluginTemplate,
  AuthTypeInfo,
  CapabilityInfo,
  ConversationStarter,
  DeclarativeAgentManifest,
  StoreTemplate,
  ValidationResult,
} from "../types.js";

// -- Schema URLs -------------------------------------------------------------

export const DECLARATIVE_AGENT_SCHEMA_URL =
  "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.0/schema.json";

export const CURRENT_VERSION = "v1.0";

export const TEAMS_MANIFEST_SCHEMA =
  "https://developer.microsoft.com/json-schemas/teams/v1.17/MicrosoftTeams.schema.json";

// -- Templates ---------------------------------------------------------------

export function getDeclarativeAgentTemplate(): DeclarativeAgentManifest {
  return {
    $schema: DECLARATIVE_AGENT_SCHEMA_URL,
    version: CURRENT_VERSION,
    name: "",
    description: "",
    instructions: "",
    conversation_starters: [],
    capabilities: [],
    actions: [],
  };
}

export function getApiPluginTemplate(): ApiPluginTemplate {
  return {
    schema_version: "v2.1",
    name_for_human: "",
    description_for_human: "",
    description_for_model: "",
    contact_email: "",
    namespace: "",
    capabilities: {
      localization: {},
      conversation_starters: [],
    },
    functions: [],
    runtimes: [
      {
        type: "OpenApi",
        auth: { type: "None" },
        spec: { url: "" },
        run_for_functions: [],
      },
    ],
  };
}

// -- Capabilities ------------------------------------------------------------

export function getCapabilitiesInfo(): Record<string, CapabilityInfo> {
  return {
    WebSearch: {
      description: "Enable web search capability for current information",
      use_cases: [
        "Current events and news",
        "Real-time information",
        "General knowledge queries",
        "Research and fact-checking",
      ],
      schema: { name: "WebSearch" },
      requirements: [],
    },
    GraphConnectors: {
      description: "Access data from Microsoft Graph connectors",
      use_cases: [
        "Enterprise search across custom data sources",
        "Third-party data integration",
        "Custom content indexing",
        "Unified data access",
      ],
      schema: {
        name: "GraphConnectors",
        connections: [{ connection_id: "your_connector_id" }],
      },
      requirements: ["Graph connector must be created and configured"],
    },
    OneDriveAndSharePoint: {
      description: "Search and access OneDrive and SharePoint content",
      use_cases: [
        "Document search and retrieval",
        "Collaborative content access",
        "File-based knowledge base",
        "Project documentation",
      ],
      schema: {
        name: "OneDriveAndSharePoint",
        items_by_url: [
          { url: "https://contoso.sharepoint.com/sites/SiteName" },
        ],
      },
      requirements: ["SharePoint site URL or OneDrive location"],
    },
    MicrosoftGraph: {
      description: "Direct Microsoft Graph API access",
      use_cases: [
        "User profile information",
        "Email and calendar access",
        "Teams data",
        "Files and sites access",
      ],
      schema: {
        name: "MicrosoftGraph",
        allowed_scopes: ["User.Read", "Mail.Read", "Calendars.Read"],
      },
      supported_scopes: [
        "User.Read",
        "User.ReadBasic.All",
        "Mail.Read",
        "Mail.ReadBasic",
        "Calendars.Read",
        "Calendars.ReadBasic",
        "Files.Read.All",
        "Sites.Read.All",
        "Chat.Read",
        "Team.ReadBasic.All",
      ],
      requirements: ["Appropriate permissions must be granted"],
    },
  };
}

// -- Capability alias map (friendly names → canonical) -----------------------

export const CAPABILITY_ALIAS_MAP: Record<string, string> = {
  web_search: "WebSearch",
  answer_questions: "WebSearch",
  graph_connectors: "GraphConnectors",
  onedrive_sharepoint: "OneDriveAndSharePoint",
  microsoft_graph: "MicrosoftGraph",
  graph_api: "MicrosoftGraph",
};

// -- Conversation-starters guidelines ----------------------------------------

export function getConversationStartersGuidelines(): Record<string, unknown> {
  return {
    best_practices: [
      "Be specific and actionable",
      "Show concrete examples of agent capabilities",
      "Use natural language as users would ask",
      "Limit to 3-5 starters for clarity",
      "Mix simple and advanced queries",
      "Cover key features of the agent",
    ],
    schema: {
      title: "Short descriptive title (max 50 chars)",
      text: "The actual prompt that will be sent to the agent",
    },
    good_examples: [
      { title: "Search tickets", text: "Show me all open high-priority tickets assigned to me" },
      { title: "Create ticket", text: "I need to report a login issue with my account" },
    ],
    bad_examples: [
      { title: "Help", text: "Help me", why_bad: "Too vague, doesn't demonstrate specific capability" },
      { title: "Start", text: "Get started", why_bad: "Generic, doesn't show what agent can do" },
    ],
  };
}

// -- Instructions guidelines -------------------------------------------------

export function getInstructionsGuidelines(): Record<string, string> {
  return {
    purpose: "Instructions define the agent's persona, behavior, and response style",
    max_length: "8000 characters",
    external_file: "Can use $[file('instructions.txt')] to reference external file",
    structure: `## Recommended Structure:

1. **Role Definition** – Who is the agent and its specialization
2. **Core Capabilities** – List specific features
3. **Behavior Guidelines** – Tone, style, when to ask clarifying questions
4. **Response Format** – How to structure answers
5. **Action Triggers** – When to use API plugins, parameter requirements
6. **Examples** – Sample queries and expected responses
7. **Constraints** – What the agent should NOT do, when to escalate`,
    example: `You are a Customer Support Agent specialized in helping users with technical issues.

## Your Role
- Assist users with product inquiries and technical problems
- Create and manage support tickets
- Provide status updates on existing issues
- Escalate complex problems to human agents when needed

## Response Guidelines
- Be professional, empathetic, and solution-focused
- Ask clarifying questions when information is incomplete
- Provide step-by-step troubleshooting when appropriate
- Always confirm ticket creation and provide ticket numbers

## When to Use Actions
- Use getTicket when users ask about specific ticket numbers
- Use createTicket when users report new issues
- Use searchTickets when users want to see their ticket history`,
  };
}

// -- Deployment checklist ----------------------------------------------------

export interface DeploymentStep {
  step: string;
  items: string[];
}

export function getDeploymentChecklist(): DeploymentStep[] {
  return [
    {
      step: "Validate Manifest",
      items: [
        "Schema URL is correct",
        "All required fields present",
        "Field length limits respected",
        "Instructions are clear and complete",
      ],
    },
    {
      step: "Test API Plugins",
      items: [
        "All endpoints are accessible",
        "Authentication works correctly",
        "OpenAPI spec is valid",
        "Response formats are correct",
        "Error handling is implemented",
      ],
    },
    {
      step: "Prepare Icons",
      items: [
        "Color icon: 192×192 PNG",
        "Outline icon: 32×32 PNG",
        "Icons follow design guidelines",
      ],
    },
    {
      step: "Create Package",
      items: [
        "All files in correct structure",
        "ZIP file created",
        "Package validated with Teams Toolkit",
      ],
    },
    {
      step: "Test in Teams",
      items: [
        "Upload custom app",
        "Test all conversation starters",
        "Verify all capabilities work",
        "Test error scenarios",
        "Verify on mobile and desktop",
      ],
    },
    {
      step: "Publish",
      items: [
        "Submit to Teams Admin Center or AppSource",
        "Provide admin documentation",
        "Include support information",
      ],
    },
  ];
}

// -- Authentication types ----------------------------------------------------

export function getAuthenticationTypes(): Record<string, AuthTypeInfo> {
  return {
    None: {
      description: "No authentication required (public APIs)",
      schema: { type: "None" },
      use_cases: ["Public APIs", "Testing", "Internal non-sensitive data"],
    },
    OAuthPluginVault: {
      description: "OAuth 2.0 authentication with credentials stored in Plugin Vault",
      schema: { type: "OAuthPluginVault", reference_id: "unique_oauth_reference" },
      use_cases: ["Production APIs", "User-delegated permissions", "Secure access"],
      best_practice: "Recommended for production deployments",
    },
    ApiKeyPluginVault: {
      description: "API Key authentication with key stored in Plugin Vault",
      schema: { type: "ApiKeyPluginVault", reference_id: "unique_apikey_reference" },
      use_cases: ["Service-to-service", "Simple authentication", "Legacy APIs"],
    },
  };
}

// -- Common patterns ---------------------------------------------------------

export function getCommonPatterns(): Record<string, AgentPattern> {
  return {
    data_retrieval: {
      name: "Data Retrieval Agent",
      description: "Search and retrieve information from enterprise systems",
      components: [
        "GraphConnectors for enterprise data",
        "OneDriveAndSharePoint for documents",
        "API plugins for external systems",
      ],
      example_use_cases: [
        "Knowledge base search",
        "Document lookup",
        "Policy information retrieval",
      ],
    },
    task_automation: {
      name: "Task Automation Agent",
      description: "Create, update, and manage tasks or tickets",
      components: [
        "API plugins with CRUD operations",
        "OAuth authentication",
        "Adaptive Cards for rich responses",
      ],
      example_use_cases: [
        "IT helpdesk ticketing",
        "Project task management",
        "Workflow automation",
      ],
    },
    notification: {
      name: "Notification Agent",
      description: "Monitor and alert on important events",
      components: [
        "API plugins for event sources",
        "Proactive messaging",
        "Graph connectors for preferences",
      ],
      example_use_cases: [
        "System monitoring alerts",
        "Business metric notifications",
        "Deadline reminders",
      ],
    },
    analysis: {
      name: "Analysis Agent",
      description: "Analyze data and provide insights",
      components: [
        "Multiple data source connectors",
        "API plugins for analytics services",
        "Rich formatting for results",
      ],
      example_use_cases: [
        "Sales analytics",
        "Performance reporting",
        "Trend analysis",
      ],
    },
  };
}

// -- OpenAPI best practices --------------------------------------------------

export function getOpenApiBestPractices(): string[] {
  return [
    "Use OpenAPI 3.0 or higher",
    "Provide detailed operation descriptions for the AI model",
    "Use clear, descriptive operationId values",
    "Include examples in schema definitions",
    "Define proper error responses (400, 401, 404, 500)",
    "Use security schemes (OAuth 2.0 recommended)",
    "Keep operations focused and single-purpose",
    "Use appropriate HTTP methods (GET for reads, POST for creates)",
    "Include parameter descriptions and constraints",
    "Version your API properly",
    "Use HTTPS endpoints only",
    "Implement proper pagination for list operations",
    "Return meaningful error messages",
    "Use standard HTTP status codes correctly",
  ];
}

// -- Validation --------------------------------------------------------------

export function validateAgentManifest(manifest: Record<string, unknown>): ValidationResult {
  const errors: string[] = [];
  const warnings: string[] = [];

  const requiredFields = ["$schema", "version", "name", "description"];
  for (const field of requiredFields) {
    if (!(field in manifest)) {
      errors.push(`Missing required field: ${field}`);
    }
  }

  if (manifest["$schema"] !== DECLARATIVE_AGENT_SCHEMA_URL) {
    errors.push(`Invalid schema URL. Expected: ${DECLARATIVE_AGENT_SCHEMA_URL}`);
  }

  if (manifest["version"] !== CURRENT_VERSION) {
    errors.push(`Invalid version. Expected: ${CURRENT_VERSION}`);
  }

  if (typeof manifest["name"] === "string" && manifest["name"].length > 100) {
    errors.push("Name exceeds 100 characters");
  }

  if (typeof manifest["description"] === "string" && manifest["description"].length > 1000) {
    errors.push("Description exceeds 1000 characters");
  }

  if (typeof manifest["instructions"] === "string") {
    const instr = manifest["instructions"] as string;
    if (!instr.startsWith("$[file(") && instr.length > 8000) {
      errors.push("Instructions exceed 8000 characters. Consider using an external file reference.");
    }
  }

  if (!("instructions" in manifest) && !("capabilities" in manifest)) {
    warnings.push("Agent has no instructions or capabilities defined");
  }

  const starters = manifest["conversation_starters"];
  if (!Array.isArray(starters) || starters.length === 0) {
    warnings.push("No conversation starters defined. Consider adding 3-5 starters.");
  }

  return { is_valid: errors.length === 0, errors, warnings };
}

// -- Agent Store catalog -----------------------------------------------------

export function getStoreTemplates(): Record<string, StoreTemplate> {
  return {
    "customer-support": {
      name: "Customer Support Agent",
      category: "Support",
      description: "AI agent for customer support with ticket management",
      author: "Microsoft",
      rating: 4.8,
      downloads: 15000,
      capabilities: ["WebSearch", "MicrosoftGraph"],
    },
    "sales-assistant": {
      name: "Sales Assistant",
      category: "Sales",
      description: "CRM integration and lead qualification agent",
      author: "Microsoft",
      rating: 4.6,
      downloads: 12000,
      capabilities: ["MicrosoftGraph", "OneDriveAndSharePoint"],
    },
    "hr-onboarding": {
      name: "HR Onboarding Agent",
      category: "HR",
      description: "Automate employee onboarding process",
      author: "Community",
      rating: 4.5,
      downloads: 8000,
      capabilities: ["MicrosoftGraph", "WebSearch"],
    },
    "data-analyst": {
      name: "Data Analyst Agent",
      category: "Analytics",
      description: "Analyze data and generate insights",
      author: "Microsoft",
      rating: 4.7,
      downloads: 10000,
      capabilities: ["MicrosoftGraph", "GraphConnectors"],
    },
    "meeting-assistant": {
      name: "Meeting Assistant",
      category: "Productivity",
      description: "Schedule meetings and manage calendars",
      author: "Community",
      rating: 4.4,
      downloads: 7500,
      capabilities: ["MicrosoftGraph"],
    },
  };
}

// -- Conversation starters per template ------------------------------------

export function getTemplateConversationStarters(
  templateId: string
): ConversationStarter[] {
  const map: Record<string, ConversationStarter[]> = {
    "customer-support": [
      { title: "Check ticket status", text: "What's the status of ticket #12345?" },
      { title: "Report issue", text: "I need to report a technical problem" },
      { title: "Get help", text: "How do I reset my password?" },
    ],
    "sales-assistant": [
      { title: "Qualify lead", text: "Tell me about the new lead from Contoso" },
      { title: "Check pipeline", text: "Show me my sales pipeline" },
      { title: "Schedule demo", text: "Schedule a product demo with the prospect" },
    ],
    "hr-onboarding": [
      { title: "Start onboarding", text: "I need to onboard a new employee" },
      { title: "Check progress", text: "What's the onboarding status for John Doe?" },
      { title: "Get checklist", text: "Show me the onboarding checklist" },
    ],
    "data-analyst": [
      { title: "Analyze data", text: "Analyze sales data for last quarter" },
      { title: "Generate report", text: "Create a revenue report" },
      { title: "Show trends", text: "What are the trends in customer acquisition?" },
    ],
    "meeting-assistant": [
      { title: "Schedule meeting", text: "Schedule a team meeting for next week" },
      { title: "Check availability", text: "When is everyone available on Friday?" },
      { title: "Meeting summary", text: "Summarize yesterday's meeting" },
    ],
  };
  return map[templateId] ?? [];
}

// -- Trigger & action maps for workflows ------------------------------------

export const TRIGGER_MAP: Record<string, Record<string, unknown>> = {
  document_uploaded: { type: "OnFileCreated", site: "SharePoint", library: "Documents" },
  email_received: { type: "OnEmailReceived", folder: "Inbox" },
  schedule: { type: "Recurrence", frequency: "Day", interval: 1 },
  form_submitted: { type: "OnFormSubmit", formId: "FORM_ID" },
};

export const ACTION_MAP: Record<string, Record<string, unknown>> = {
  send_notification: { type: "SendEmail", to: "user@example.com", subject: "Notification", body: "Action completed" },
  request_approval: { type: "StartApproval", approvers: ["manager@example.com"], title: "Approval Required" },
  create_task: { type: "CreateTask", taskList: "Tasks", title: "New Task" },
  update_database: { type: "UpdateRow", table: "Data", columns: {} },
  send_teams_message: { type: "PostToChannel", channel: "General", message: "Update from workflow" },
};
