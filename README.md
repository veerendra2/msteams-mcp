# Microsoft Teams MCP Server

MCP server for Microsoft Teams — search, message, and manage chats with browser-session auth; no Teams app approval or Azure AD registration.

Clone the repo, build it, point your AI tool at `dist/index.js`. Talks to the same Teams APIs the web app uses (Substrate, chatsvc, CSA). Auth is your normal Teams browser login; access matches what your account can already do.

## Prerequisites

- Node.js 18+
- Microsoft account with Teams access
- Chrome (macOS/Linux) or Edge (Windows)

## Setup

```bash
git clone https://github.com/hickeroar/msteams-mcp.git
cd msteams-mcp
npm install
npm run build
```

Register it as a **local stdio MCP server** in whatever AI tool you use (Cursor, Claude Desktop, Claude Code, Augment, Windsurf, etc.). The shape is always the same: run `node` with the absolute path to `dist/index.js`.

```json
{
  "mcpServers": {
    "MS Teams": {
      "command": "node",
      "args": ["/absolute/path/to/msteams-mcp/dist/index.js"]
    }
  }
}
```

Where that config lives depends on the client (examples):

| Client | Typical config location |
|--------|-------------------------|
| Cursor | `~/.cursor/mcp.json` (or project `.cursor/mcp.json`) |
| Claude Desktop | Claude → Settings → Developer → Edit Config (`claude_desktop_config.json`) |
| Claude Code | MCP settings / `.mcp.json` as documented for your version |
| Augment | Augment MCP / tools settings (add a local stdio server) |

Exact UI labels vary; look for “MCP servers”, “Model Context Protocol”, or “custom tools”, then add a **command + args** (stdio) entry like above. Use an absolute path.

Restart or reload MCP/tools in that client after changing config or rebuilding.

### First login

Ask the assistant (in Cursor, Claude, Augment, etc.) to call `teams_login` (soft `{}` first; use `forceNew: true` if the session is dead). A browser opens for Microsoft/SSO sign-in. Session files are saved under:

- **macOS/Linux:** `~/.teams-mcp-server/`
- **Windows:** `%APPDATA%\teams-mcp-server\`

After that, tools use cached tokens. Tokens refresh automatically when possible; call `teams_login` again if auth fails.

### After pulling code changes

```bash
npm install
npm run build
```

Then restart/reload the Teams MCP server in your AI client so it loads the new `dist/`.

## Useful commands

```bash
npm run build        # compile to dist/
npm run typecheck    # tsc --noEmit
npm start            # run compiled MCP server (stdio)
npm run cli          # CLI with the same tools (no build required)
npm run cli -- status
npm run cli -- login
npm run cli -- find_channel --query "general"
```

The CLI shares the same session as the MCP server and is handy for debugging outside any AI client.

## Available tools

### Search & discovery

| Tool | Description |
|------|-------------|
| `teams_search` | Search Teams messages with operators (`from:`, `sent:`, `in:`, etc.) |
| `teams_search_email` | Search emails in your mailbox |
| `teams_list_chats` | List recent conversations with last-message preview |
| `teams_get_message` | Get a single message by ID (includes reactions) |
| `teams_get_thread` | Get messages from a conversation/thread; `since`, `threadRootId`, `fromUrl` |
| `teams_find_channel` | Find channels by name (membership list; works without Substrate) |
| `teams_get_activity` | Activity feed (mentions, reactions, replies, notifications) |

### Messaging

| Tool | Description |
|------|-------------|
| `teams_send_message` | Send a message (markdown); thread reply, subject, schedule, content type |
| `teams_wait_for_reply` | Block until a new message arrives (capped ~110s) |
| `teams_edit_message` | Edit one of your own messages |
| `teams_delete_message` | Soft-delete one of your own messages |

### People & contacts

| Tool | Description |
|------|-------------|
| `teams_get_me` | Current user profile |
| `teams_search_people` | Search people by name or email |
| `teams_get_frequent_contacts` | Frequently contacted people |
| `teams_get_person` | Resolve MRI(s) to full profiles |
| `teams_get_chat` | Conversation ID for a 1:1 chat |
| `teams_create_group_chat` | Create a group chat (2+ others) |

### Organisation

| Tool | Description |
|------|-------------|
| `teams_get_favorites` | Pinned/favourite conversations |
| `teams_add_favorite` | Pin a conversation |
| `teams_remove_favorite` | Unpin a conversation |
| `teams_save_message` | Bookmark a message |
| `teams_unsave_message` | Remove bookmark |
| `teams_get_saved_messages` | List saved messages |
| `teams_get_followed_threads` | List followed threads |
| `teams_get_unread` | Unread counts (aggregate or per conversation) |
| `teams_mark_read` | Mark a conversation read up to a message |
| `teams_get_tags` | Channel tags for a team (`teamId` from `teams_find_channel`) |

### Reactions

| Tool | Description |
|------|-------------|
| `teams_search_emoji` | Search standard + custom org emojis |
| `teams_add_reaction` | Add a reaction to a message |
| `teams_remove_reaction` | Remove a reaction |

Quick reactions without search: `like`, `heart`, `laugh`, `surprised`, `sad`, `angry`.

### Calendar & meetings

| Tool | Description |
|------|-------------|
| `teams_get_meetings` | Calendar meetings (default: next 7 days) |
| `teams_get_transcript` | Meeting transcript (`threadId` from `teams_get_meetings`) |

### Files

| Tool | Description |
|------|-------------|
| `teams_get_shared_files` | Files and links shared in a conversation |

### Session

| Tool | Description |
|------|-------------|
| `teams_login` | Manual login (opens browser) |
| `teams_status` | Auth / session status |

Channel reads usually work via `teams_find_channel` → `teams_get_thread` even when Substrate search (`teams_search`) is unavailable. Architecture notes: [AGENTS.md](AGENTS.md). API details: [docs/API-REFERENCE.md](docs/API-REFERENCE.md).

## Limitations

- Browser login required once (and again if the session fully expires)
- Undocumented Microsoft APIs; they can change
- Search is full-text only; use `teams_get_thread` for full thread context
- Edit/delete only your own messages

## License

MIT
