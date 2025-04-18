# Yahoo Mail MCP Server

A Model Context Protocol (MCP) server that integrates with Yahoo Mail through IMAP. By using this server, you can easily access and manage your Yahoo Mail account through an AI assistant.

## Features

- List email folders in your Yahoo Mail account
- Search for emails with flexible query options
- Read email content with proper decoding
- Check unread email count in folders
- Mark emails as read or unread
- Move emails between folders
- Delete emails

## Prerequisites

- Python 3.13 or higher
- Yahoo Mail account with IMAP access enabled
- MCP-compatible client (such as Claude Desktop)

## Installation

1. Clone this repository:
  ```bash
  git clone --filter=blob:none --no-checkout https://github.com/KeisukeYamazaki/mcp-servers.git
  cd mcp-servers
  
  # If using sparse-checkout for the first time
  git sparse-checkout init --cone
  git sparse-checkout set ymail
  git checkout
  
  # If you've already checked out other directories, use this instead
  # git sparse-checkout add ymail
  # git checkout
  ```

## Configuration

Set the following environment variables with your Yahoo Mail credentials:

```bash
export YAHOO_EMAIL="your.email@yahoo.com"
export YAHOO_PASSWORD="your-password"
```

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
      "name": "ymail",
      "command": "uv",
      "args": [
        "--directory",
        "path/to/ymail",
        "run",
        "main.py"
      ],
      "env": {
        "YAHOO_EMAIL": "your.email@yahoo.com",
        "YAHOO_PASSWORD": "your-password"
      }
    }
  ]
}
```

3. Restart Claude Desktop.

## Available Tools

| Tool Name          | Description                                 |
| ------------------ | ------------------------------------------- |
| `list_folders`     | List all available email folders            |
| `search_emails`    | Search for emails in a specified folder     |
| `read_email`       | Read a specific email by ID                 |
| `get_unread_count` | Get the number of unread emails in a folder |
| `mark_as_read`     | Mark an email as read                       |
| `mark_as_unread`   | Mark an email as unread                     |
| `move_email`       | Move an email from one folder to another    |
| `delete_email`     | Delete an email (moves to Trash)            |

## Example Conversations

### Checking Unread Emails

You: "Do I have any unread emails?"

Claude: "Let me check your unread emails."
*Uses `get_unread_count` tool*
"You have 5 unread emails in your inbox."

### Reading a Specific Email

You: "Can you show me my most recent emails and then read the one from Amazon?"

Claude: "Let me find your recent emails."
*Uses `search_emails` tool*
"Here are your most recent emails:
1. Amazon Order Confirmation (ID: 12345)
2. LinkedIn Notification (ID: 12346)
3. Newsletter Subscription (ID: 12347)

Let me get the details of the Amazon email for you."
*Uses `read_email` tool with ID 12345*
"Here's the email from Amazon:
Subject: Your Amazon Order #123-4567890
From: order-confirmation@amazon.com
Date: June 1, 2023
Body: Thank you for your order. Your package will arrive..."

## License

[MIT License](LICENSE)
