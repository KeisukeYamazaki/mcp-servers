# Google Spreadsheet MCP Server

A Model Context Protocol (MCP) server that provides tools for interacting with Google Spreadsheets. This server enables AI assistants to create, read, update, and manage Google Spreadsheets through the MCP standard.

## Features

- Create new Google Spreadsheets
- Copy existing spreadsheets with optional destination folder
- Rename spreadsheets
- List all sheets within a spreadsheet
- Create a copy of a sheet within a spreadsheet
- Rename sheets
- Get data from specific sheets and ranges
- Add rows and columns to sheets
- Update single cells or ranges of cells

## Prerequisites

- Python 3.13 or higher
- Google API credentials with appropriate permissions
- MCP-compatible client (such as Claude Desktop)

## Installation

1. Clone this repository:
  ```bash
  git clone --filter=blob:none --no-checkout https://github.com/KeisukeYamazaki/mcp-servers.git
  cd mcp-servers
  
  # If using sparse-checkout for the first time
  git sparse-checkout init --cone
  git sparse-checkout set gsheet
  git checkout
  
  # If you've already checked out other directories, use this instead
  # git sparse-checkout add gsheet
  # git checkout
  ```

2. Configure Google API credentials:
  - Create a project in the [Google Cloud Console](https://console.cloud.google.com/)
  - Enable the Google Drive API and Google Sheets API
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
      "name": "gsheet",
      "command": "uv",
      "args": [
        "--directory",
        "path/to/gsheet",
        "run",
        "main.py"
      ]
    }
  ]
}
```

3. Restart Claude Desktop.

## Available Tools

| Tool Name            | Description                                |
| -------------------- | ------------------------------------------ |
| `list_spreadsheets`  | List all spreadsheets or those in a folder |
| `copy_spreadsheet`   | Copy a spreadsheet with a new name         |
| `create_spreadsheet` | Create a new spreadsheet                   |
| `rename_spreadsheet` | Rename an existing spreadsheet             |
| `list_sheets`        | List all sheets within a spreadsheet       |
| `copy_sheet`         | Copy a sheet within a spreadsheet          |
| `rename_sheet`       | Rename a sheet within a spreadsheet        |
| `get_sheet_data`     | Get data from a specific sheet/range       |
| `add_rows`           | Add rows to a sheet                        |
| `add_columns`        | Add columns to a sheet                     |
| `update_cell`        | Update a single cell in a sheet            |
| `update_cells`       | Update multiple cells in a sheet           |

## Example Conversations

### Creating and updating a spreadsheet

You: "Can you create a new spreadsheet titled 'Monthly Budget' and add column headers for 'Category', 'Amount', and 'Notes'?"

Claude: "I'll create a new spreadsheet titled 'Monthly Budget' for you with those headers."
*Uses `create_spreadsheet` tool*
"I've created the spreadsheet. Now I'll add the column headers."
*Uses `update_cells` tool*
"Your spreadsheet 'Monthly Budget' is ready with headers for 'Category', 'Amount', and 'Notes'."

### Reading spreadsheet data

You: "What data is in my 'Q1 Sales' spreadsheet, on the 'January' sheet?"

Claude: "Let me check that spreadsheet for you."
*Uses `get_sheet_data` tool*
"Here's the data from your 'January' sheet in the 'Q1 Sales' spreadsheet:
- Row 1: ['Product', 'Units Sold', 'Revenue']
- Row 2: ['Product A', '42', '$4,200']
- Row 3: ['Product B', '28', '$1,400']
..."

## License

[MIT License](LICENSE)
