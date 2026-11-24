// ---------------------------------------------------------------------------
// Shared types for the M365 Toolkit MCP Server
// ---------------------------------------------------------------------------

export enum ResponseFormat {
  MARKDOWN = "markdown",
  JSON = "json",
}

// -- Declarative Agent types ------------------------------------------------

export interface ConversationStarter {
  title: string;
  text: string;
}

export interface CapabilityConfig {
  name: string;
  [key: string]: unknown;
}

export interface DeclarativeAgentManifest {
  $schema: string;
  version: string;
  name: string;
  description: string;
  instructions: string;
  conversation_starters: ConversationStarter[];
  capabilities: CapabilityConfig[];
  actions: unknown[];
}

export interface ValidationResult {
  is_valid: boolean;
  errors: string[];
  warnings: string[];
}

// -- Capability info --------------------------------------------------------

export interface CapabilityInfo {
  description: string;
  use_cases: string[];
  schema: CapabilityConfig;
  requirements: string[];
  supported_scopes?: string[];
}

// -- Adaptive Card types ----------------------------------------------------

export interface AdaptiveCard {
  $schema: string;
  type: "AdaptiveCard";
  version: string;
  body: unknown[];
  actions?: unknown[];
}

// -- Workflow types ----------------------------------------------------------

export interface WorkflowConfig {
  name: string;
  description: string;
  triggers: Record<string, unknown>[];
  actions: Record<string, unknown>[];
  enabled: boolean;
}

// -- Teams Bot types --------------------------------------------------------

export interface TeamsBotManifest {
  $schema: string;
  manifestVersion: string;
  version: string;
  id: string;
  packageName: string;
  developer: {
    name: string;
    websiteUrl: string;
    privacyUrl: string;
    termsOfUseUrl: string;
  };
  icons: { color: string; outline: string };
  name: { short: string; full: string };
  description: { short: string; full: string };
  accentColor: string;
  bots: unknown[];
  permissions: string[];
  validDomains: string[];
}

// -- Template Store types ---------------------------------------------------

export interface StoreTemplate {
  name: string;
  category: string;
  description: string;
  author: string;
  rating: number;
  downloads: number;
  capabilities: string[];
}

// -- API Plugin types -------------------------------------------------------

export interface ApiPluginTemplate {
  schema_version: string;
  name_for_human: string;
  description_for_human: string;
  description_for_model: string;
  contact_email: string;
  namespace: string;
  capabilities: { localization: Record<string, unknown>; conversation_starters: unknown[] };
  functions: unknown[];
  runtimes: unknown[];
}

// -- Auth types -------------------------------------------------------------

export interface AuthTypeInfo {
  description: string;
  schema: Record<string, string>;
  use_cases: string[];
  best_practice?: string;
}

// -- Pattern types ----------------------------------------------------------

export interface AgentPattern {
  name: string;
  description: string;
  components: string[];
  example_use_cases: string[];
}
