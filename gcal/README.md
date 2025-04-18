# Google Calendar MCP Server

A Model Context Protocol (MCP) server that integrates with the Google Calendar API. By using this server, you can easily perform operations such as checking and updating Google Calendar schedules through an AI assistant.

## Features

- Retrieve and display calendar events
- Get details of specific events
- Create new calendar events
- Update existing calendar events
- Delete calendar events
- Search calendar events by keyword
- Get list of today's events

## Prerequisites

- Python 3.8 or higher
- Google API credentials with appropriate permissions
- MCP-compatible client (such as Claude Desktop)

## Installation

1. Clone this repository:
  ```bash
  git clone --filter=blob:none --no-checkout https://github.com/KeisukeYamazaki/mcp-servers.git
  cd mcp-servers
  
  # If using sparse-checkout for the first time
  git sparse-checkout init --cone
  git sparse-checkout set gcal
  git checkout
  
  # If you've already checked out other directories, use this instead
  # git sparse-checkout add gcal
  # git checkout
  ```

2. Configure Google API credentials:
  - Create a project in the [Google Cloud Console](https://console.cloud.google.com/)
  - Enable the Google Calendar API
  - Create OAuth credentials and download the credentials JSON file
  - Place the credentials file in the project directory as `credentials.json`

3. Install the required packages:
   ```bash
   uv add "mcp[cli]"
   uv add google-api-python-client google-auth-httplib2 google-auth-oauthlib
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
      "name": "gcal",
      "command": "uv",
      "args": [
        "--directory",
        "path/to/gcal",
        "run",
        "main.py"
      ]
    }
  ]
}
```

3. Restart Claude Desktop.

## Available Tools

| Tool Name              | Description                       |
| ---------------------- | --------------------------------- |
| `list_upcoming_events` | Retrieve upcoming calendar events |
| `get_event_details`    | Get details of a specific event   |
| `create_event`         | Create a new calendar event       |
| `update_event`         | Update an existing calendar event |
| `delete_event`         | Delete a calendar event           |
| `search_events`        | Search calendar events by keyword |

## Example Conversations

### Creating an Event

You: "Can you create a project meeting from 10 AM to 11 AM on June 1st?"

Claude: "I'll create a project meeting for June 1st from 10 AM to 11 AM."
*Uses `create_event` tool*
"I've created the event. Your Project Meeting is now scheduled for June 1st from 10:00 to 11:00."

### Checking Upcoming Events

You: "What events do I have in the next two weeks?"

Claude: "Let me check your upcoming events for the next two weeks."
*Uses `list_upcoming_events` tool*
"Here are your events for the next two weeks:
- June 1, 10:00-11:00: Project Meeting
- June 3, 13:00-14:00: Team Lunch
- June 8, 15:00-16:30: Client Presentation
..."

## License

[MIT License](LICENSE)
