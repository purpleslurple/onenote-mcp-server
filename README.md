# OneNote MCP Server

A complete, robust Model Context Protocol (MCP) server for Microsoft OneNote integration with Claude Desktop. Access your entire OneNote knowledge base through natural language queries.

## 🎯 What This Does

Transform your OneNote notebooks into an AI-accessible knowledge base:
- **List all your notebooks, sections, and pages**
- **Read page content** for analysis and search
- **Natural language queries** like "Show me my DevOps notes" or "Find pages about project planning"
- **Secure OAuth authentication** with Microsoft Graph API
- **Bulletproof error handling** with detailed debugging

## ✨ Why This Implementation

Unlike other OneNote MCP servers, this one:
- ✅ **Actually works** - tested extensively with real OneNote data
- ✅ **Complete functionality** - all core OneNote operations implemented
- ✅ **Robust authentication** - two-step device flow that handles edge cases
- ✅ **Production ready** - proper error handling and logging
- ✅ **Easy setup** - detailed instructions for non-technical users

## 🚀 Quick Start

### Prerequisites
- Python 3.10+ 
- [uv package manager](https://docs.astral.sh/uv/getting-started/installation/) (recommended) or pip
- Claude Desktop
- Microsoft Azure account (free)

### 1. Install uv (if you don't have it)
```bash
# macOS/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh

# or with Homebrew
brew install uv
```

### 2. Clone and Setup
```bash
git clone https://github.com/yourusername/onenote-mcp-server.git
cd onenote-mcp-server

# Create virtual environment and install dependencies
uv sync
```

### 3. Azure App Registration
You need to create an Azure app to access OneNote. **Don't worry, it's free and takes 5 minutes:**

1. Go to [Azure Portal](https://portal.azure.com) (sign in with your Microsoft account)
2. Navigate to **Azure Active Directory** → **App registrations** → **New registration**
3. Fill out the form:
   - **Name**: "OneNote MCP Server" (or whatever you like)
   - **Supported account types**: "Accounts in any organizational directory and personal Microsoft accounts"
   - **Redirect URI**: Select "Public client/native" and enter: `https://login.microsoftonline.com/common/oauth2/nativeclient`
4. Click **Register**
5. Copy the **Application (client) ID** - you'll need this!

### 4. Add Permissions
Still in your Azure app:
1. Go to **API permissions** → **Add a permission**
2. Select **Microsoft Graph** → **Delegated permissions**
3. Add these permissions:
   - `Notes.Read` - Read OneNote notebooks
   - `Notes.ReadWrite` - Create/modify OneNote content (optional but recommended)
   - `User.Read` - Read user profile
4. Click **Grant admin consent** (the button at the top)

### 5. Configure Claude Desktop
Edit your Claude Desktop config file:

**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows**: `%APPDATA%\\Claude\\claude_desktop_config.json`

Add this configuration (replace `YOUR_CLIENT_ID` with your actual Azure client ID):

```json
{
  "mcpServers": {
    "onenote": {
      "command": "uv",
      "args": [
        "--directory", "/FULL/PATH/TO/onenote-mcp-server",
        "run", "python", "onenote_mcp_server.py"
      ],
      "env": {
        "AZURE_CLIENT_ID": "YOUR_CLIENT_ID_HERE"
      }
    }
  }
}
```

**Important**: Replace `/FULL/PATH/TO/onenote-mcp-server` with the complete path to where you cloned this repo.

### 6. Restart Claude Desktop
Completely quit and restart Claude Desktop. You should see OneNote tools in the 🔨 menu.

## 🔐 First Time Authentication

1. In Claude Desktop, say: **"Start OneNote authentication"**
2. Claude will give you a URL and code
3. Visit the URL in your browser, enter the code, and sign in
4. Come back to Claude and say: **"Complete OneNote authentication"**
5. You're ready to go!

## 📖 Usage Examples

Once authenticated, try these commands in Claude Desktop:

```
List my OneNote notebooks
Show me sections in my Work notebook  
What pages are in my Ideas section?
Read the content of my "Project Plan" page
```

## 🛠 Troubleshooting

### "No tools available" in Claude Desktop
- Make sure you restarted Claude Desktop after config changes
- Check that the path in your config is correct (use full absolute path)
- Verify uv is installed: `uv --version`

### Authentication fails
- Double-check your Azure client ID is correct
- Make sure you granted admin consent for all permissions
- Try deleting any cached tokens and re-authenticating

### "Command not found" errors
- Make sure uv is in your PATH
- Alternative: replace `"uv"` with `"python"` in the config and use the full path to your Python interpreter

### Permission denied errors
- Check the file permissions in your project directory
- Make sure Claude Desktop can read the files

## 🏗 Development

### Project Structure
```
onenote-mcp-server/
├── onenote_mcp_server.py      # Main server implementation
├── pyproject.toml             # Dependencies and metadata  
├── README.md                  # This file
├── LICENSE                    # MIT License
└── .gitignore                 # Git ignore rules
```

### Key Features
- **Two-step authentication**: Handles device code flow properly
- **Complete Graph API integration**: All OneNote operations supported
- **Robust error handling**: Detailed logging and graceful failures
- **FastMCP framework**: Clean, maintainable code structure
- **Environment variable configuration**: Secure credential handling

### Adding New Features
The server is built with FastMCP, making it easy to add new tools:

```python
@mcp.tool()
async def your_new_tool(param: str) -> str:
    """Description of what your tool does."""
    # Your implementation here
    return result
```

## 🤝 Contributing

Contributions welcome! Please:
1. Fork the repo
2. Create a feature branch
3. Add tests for new functionality  
4. Submit a pull request

## 📄 License

MIT License - see LICENSE file for details.

## 🙏 Acknowledgments

- Built with [FastMCP](https://github.com/jlowin/fastmcp) framework
- Uses Microsoft Graph API for OneNote access
- Inspired by the amazing potential of AI + personal knowledge bases

## ⚠️ Important Notes

- This server only reads/writes data you already have access to
- Your Azure app credentials stay on your machine
- All authentication happens directly between you and Microsoft
- No data is sent to third parties

---

**Built with ❤️ for the Claude + OneNote community**

*Turn your OneNote into an AI-accessible knowledge base!*
