# Google Document MCP Server

A Model Context Protocol (MCP) server that provides tools for interacting with Google Documents. This server enables AI assistants to create, read, update, and manage Google Documents through the MCP standard.

## Features

- Create new Google Documents with custom content
- Read content from existing documents
- Update document content (append or replace)
- List documents in specific folders
- Create and manage folders
- Move documents between folders
- Delete documents
- Rename documents and folders
- Copy documents and folders

## Prerequisites

- Python 3.13 or higher
- Google API credentials with appropriate permissions
- MCP-compatible client (such as Claude Desktop)

## Installation

1. Clone this repository:
  ```bash
  git clone --filter=blob:none --no-checkout https://github.com/KeisukeYamazaki/mcp-servers.git
  cd mcp-servers
  git sparse-checkout set gdoc
  git checkout
  ```

2. Configure Google API credentials:
  - Create a project in the [Google Cloud Console](https://console.cloud.google.com/)
  - Enable the Google Drive API and Google Docs API
  - Create OAuth credentials and download the credentials JSON file
  - Place the credentials file in the project directory as `credentials.json`

## Usage

### Using with Claude Desktop

1. Edit your Claude Desktop configuration:
  - Mac: `~/Library/Application Support/Claude/claude_desktop_config.json`
  - Windows: `%APPDATA%\Claude\claude_desktop_config.json`

2. Add this server to the configuration:
```json
{
  "mcpServers": [
    {
      "name": "gdoc",
      "command": "uv",
      "args": [
        "--directory",
        "path/to/gdoc",
        "run",
        "main.py"
      ]
    }
  ]
}
```

3. Restart Claude Desktop.

## Available Tools

| Tool Name           | Description                                |
| ------------------- | ------------------------------------------ |
| `create_document`   | Create a new Google Document               |
| `read_document`     | Read the content of a document             |
| `update_document`   | Append or replace content in a document    |
| `list_documents`    | List documents in a folder                 |
| `list_folders`      | List folders                               |
| `create_folder`     | Create a new folder                        |
| `delete_document`   | Delete a document                          |
| `move_document`     | Move a document to another folder          |
| `rename_document`   | Rename a document                          |
| `rename_folder`     | Rename a folder                            |
| `copy_document`     | Create a copy of a document                |
| `copy_folder`       | Create a copy of a folder and its contents |
| `lock_document`     | Lock a document (make it read-only)        |
| `unlock_document`   | Unlock a document                          |
| `add_bulleted_list` | Add a bulleted list to a document          |
| `add_numbered_list` | Add a numbered list to a document          |
| `add_table`         | Add a table to a document                  |

## Example Conversations

### Creating and updating a document

You: "Can you create a new document titled 'Meeting Notes' and add today's date as a heading?"

Claude: "I'll create a new document titled 'Meeting Notes' for you with today's date as a heading."
*Uses `create_document` tool*
"I've created the document. Here's the link: [Meeting Notes](https://docs.google.com/document/d/abc123...)"

### Reading document content

You: "What's in my 'Project Roadmap' document?"

Claude: "Let me check that document for you."
*Uses `read_document` tool*
"Your 'Project Roadmap' document contains the following sections:
- Project Overview
- Timeline
- Milestones
- Resources
...

## License

[MIT License](LICENSE)
