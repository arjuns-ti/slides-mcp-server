# Slides MCP Server

A FastMCP server for verifying Google Slides files in Google Drive using the Model Context Protocol (MCP).

## Features

- **Presentation Management**: Read, edit, and manage Google Slides presentations
- **Text Editing**: Update text with full formatting control (bold, colors, fonts, alignment, etc.)
- **Element Management**: Add images, tables, charts, and bullets to slides
- **Slide Operations**: Duplicate slides and manage presentation structure
- **OAuth Authentication**: Secure Google API access with presentations read/write permissions
- **Logging**: Optional file-based logging system
- **No Console Output**: All logging goes to `logs.txt` for debugging

## Prerequisites

1. **Google Cloud Project Setup**:
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select an existing one
   - Enable the Google Drive API and Google Slides API
   - Create OAuth 2.0 credentials (Desktop app)
   - Configure OAuth consent screen with the following scopes:
     - `https://www.googleapis.com/auth/drive.readonly`
     - `https://www.googleapis.com/auth/drive.file`
     - `https://www.googleapis.com/auth/presentations`
   - Download the credentials JSON file

2. **UV Package Manager**:
   - Install [uv](https://github.com/astral-sh/uv) if not already installed

## Installation

```bash
# Install dependencies
uv sync
```

## Configuration

1. Copy the example environment file:
```bash
cp env.example .env
```

2. Edit `.env` and configure:
   - `OAUTH_CLIENT_SECRET`: Path to your Google OAuth credentials JSON file
   - `OAUTH_CLIENT_TOKEN`: Path where the OAuth token will be saved
   - `ENABLE_LOGGING`: Set to `true` to enable logging to `logs.txt`

3. Place your Google OAuth credentials file in the location specified by `OAUTH_CLIENT_SECRET`

## Usage

### Running the Server

Run the server in stdio mode:

```bash
uv run slides-mcp
```

### Development with Inspector

Run the server with the FastMCP inspector for testing and debugging:

```bash
uv run fastmcp dev src/server.py
```

This will start an interactive inspector where you can:
- Test tools
- Browse resources
- Try prompts
- View server capabilities

## Project Structure

```
slides-mcp-server/
├── src/
│   ├── __init__.py
│   └── server.py       # Main FastMCP server implementation
├── pyproject.toml      # Project configuration and dependencies
└── README.md
```

## Available Tools

### `get_slides_deck(file_id: str)`

Gets the contents of a Google Slides presentation by its file ID.

**Parameters:**
- `file_id` (str): The Google Drive file ID to check

**Returns:**
- `success` (bool): Whether the operation succeeded
- `is_slides` (bool): Whether the file is a Google Slides presentation
- `presentation` (object): Presentation metadata (id, title, locale, slide_count, revision_id)
- `file` (object): File metadata (id, name, mimeType, timestamps, webViewLink)
- `slides` (array): Array of slide objects with content
  - Each slide contains: slide_number, object_id, and elements array
  - Elements can be: shapes (with text content), tables, images, videos
- `error` (str): Error message if an error occurred

**Example (Success - Google Slides):**
```json
{
  "success": true,
  "is_slides": true,
  "presentation": {
    "id": "1ABC...",
    "title": "My Presentation",
    "locale": "en",
    "slide_count": 3,
    "revision_id": "ALm..."
  },
  "file": {
    "id": "1ABC...",
    "name": "My Presentation",
    "mimeType": "application/vnd.google-apps.presentation",
    "webViewLink": "https://docs.google.com/presentation/d/1ABC.../edit"
  },
  "slides": [
    {
      "slide_number": 1,
      "object_id": "slide_id_1",
      "elements": [
        {
          "object_id": "shape1",
          "type": "shape",
          "content": "Welcome to the presentation"
        },
        {
          "object_id": "img1",
          "type": "image",
          "content": "https://..."
        }
      ]
    }
  ]
}
```

**Example (Not a Slides file):**
```json
{
  "exists": true,
  "is_slides": false,
  "message": "File exists but is not a Google Slides file (type: application/pdf)"
}
```

**Example (File not found):**
```json
{
  "exists": false,
  "is_slides": false,
  "message": "File with ID '1ABC...' not found in Google Drive"
}
```

## Development

The server uses stdio transport for communication, making it compatible with MCP clients like Claude Desktop.

## Logging

When `ENABLE_LOGGING=true` is set in your `.env` file, the server will log all operations to `logs.txt` including:
- Authentication attempts
- File search queries
- Results and errors
- Timestamps for all operations

## Authentication Flow

On first run, the server will:
1. Check for an existing token at `OAUTH_CLIENT_TOKEN`
2. If not found:
   - Log the OAuth authorization URL to `logs.txt` (when `ENABLE_LOGGING=true`)
   - Open a browser for OAuth authentication
   - Save the token to the path specified in `OAUTH_CLIENT_TOKEN`
3. Automatically refresh expired tokens
4. All authentication steps are logged to `logs.txt` for debugging

**Note**: Check `logs.txt` for the authorization URL if you need to manually copy it or if the browser doesn't open automatically.

## Security Notes

- Never commit your `.env` file, `credentials.json`, or `token.json` to version control
- The OAuth token provides:
  - Read-only access to Google Drive metadata
  - Create and manage files created or opened by the app
  - Full read/write access to Google Slides presentations
- Logs may contain sensitive information - keep `logs.txt` secure
- Review and approve OAuth consent screen to control which presentations the server can access

## License

Add your license here.

