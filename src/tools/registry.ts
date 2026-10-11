/**
 * Tool registry - centralised tool management.
 * 
 * All tools are registered here and can be looked up by name.
 */

import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { z } from 'zod';
import type { ToolContext, ToolResult } from './index.js';
import { ErrorCode, createError } from '../types/errors.js';

/** Type-erased tool entry for the registry — avoids generic variance issues. */
interface RegistryEntry {
  definition: Tool;
  schema: z.ZodTypeAny;
  handler: (input: z.output<z.ZodTypeAny>, ctx: ToolContext) => Promise<ToolResult>;
}

import { searchTools } from './search-tools.js';
import { messageTools } from './message-tools.js';
import { peopleTools } from './people-tools.js';
import { authTools } from './auth-tools.js';
import { meetingTools } from './meeting-tools.js';
import { fileTools } from './file-tools.js';
import { tagTools } from './tag-tools.js';

// ─────────────────────────────────────────────────────────────────────────────
// Registry
// ─────────────────────────────────────────────────────────────────────────────

/** All registered tools (cast to base type for registry). */
const allTools: RegistryEntry[] = [
  ...searchTools,
  ...messageTools,
  ...peopleTools,
  ...authTools,
  ...meetingTools,
  ...fileTools,
  ...tagTools,
];

/** Lookup map for tools by name. */
const toolsByName = new Map<string, RegistryEntry>(
  allTools.map(tool => [tool.definition.name, tool])
);

/** Explicit allowlist so newly added tools stay hidden until reviewed. */
const READ_ONLY_TOOL_NAMES = new Set([
  'teams_status',
  'teams_search',
  'teams_search_email',
  'teams_list_chats',
  'teams_get_message',
  'teams_get_thread',
  'teams_find_channel',
  'teams_get_activity',
  'teams_wait_for_reply',
  'teams_get_me',
  'teams_search_people',
  'teams_get_frequent_contacts',
  'teams_get_person',
  'teams_get_chat',
  'teams_get_favorites',
  'teams_get_unread',
  'teams_search_emoji',
  'teams_get_saved_messages',
  'teams_get_followed_threads',
  'teams_get_meetings',
  'teams_get_transcript',
  'teams_get_shared_files',
  'teams_get_tags',
]);

const readOnlyEnabled = process.env.MSTEAMS_READ_ONLY === 'true';

function isToolAvailable(name: string): boolean {
  return !readOnlyEnabled || READ_ONLY_TOOL_NAMES.has(name);
}

// ─────────────────────────────────────────────────────────────────────────────
// Public API
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets all tool definitions for MCP ListTools.
 */
export function getToolDefinitions(): Tool[] {
  return allTools
    .filter(tool => isToolAvailable(tool.definition.name))
    .map(tool => tool.definition);
}

/**
 * Invokes a tool by name with the given arguments and context.
 */
export async function invokeTool(
  name: string,
  args: unknown,
  ctx: ToolContext
): Promise<ToolResult> {
  const tool = toolsByName.get(name);
  
  if (!tool || !isToolAvailable(name)) {
    return {
      success: false,
      error: createError(
        ErrorCode.INVALID_INPUT,
        tool ? `Tool '${name}' is disabled in read-only mode` : `Unknown tool: ${name}`,
        { retryable: false }
      ),
    };
  }

  // Validate input
  const parseResult = tool.schema.safeParse(args);
  if (!parseResult.success) {
    return {
      success: false,
      error: createError(ErrorCode.INVALID_INPUT, `Invalid input: ${parseResult.error.message}`, { retryable: false }),
    };
  }

  if (readOnlyEnabled && name === 'teams_get_thread' && parseResult.data.markRead) {
    return {
      success: false,
      error: createError(
        ErrorCode.INVALID_INPUT,
        "'markRead' is disabled in read-only mode",
        { retryable: false }
      ),
    };
  }

  // Invoke handler
  return tool.handler(parseResult.data, ctx);
}
