#!/usr/bin/env python3
"""
OneNote MCP Server

A Model Context Protocol server for Microsoft OneNote integration.
This allows Claude Desktop to read and interact with OneNote notebooks.
"""

import os
import asyncio
import json
import logging
from typing import List, Dict, Any, Optional
from pathlib import Path
import time
from msal import ConfidentialClientApplication, PublicClientApplication
import httpx
from fastmcp import FastMCP

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize FastMCP instance
mcp = FastMCP("OneNote MCP Server")

# Microsoft Graph API constants
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
SCOPES = [
    "https://graph.microsoft.com/Notes.Read",
    "https://graph.microsoft.com/Notes.ReadWrite",
    "https://graph.microsoft.com/User.Read"
]

# Token cache configuration
TOKEN_CACHE_ENABLED = os.getenv("ONENOTE_CACHE_TOKENS", "true").lower() in ("true", "1", "yes")
TOKEN_CACHE_FILE = Path.home() / ".onenote_mcp_tokens.json"

# Global variables for authentication
access_token: Optional[str] = None
refresh_token: Optional[str] = None
token_expires_at: Optional[float] = None
msal_app: Optional[PublicClientApplication] = None

def get_client_id() -> str:
    """Get the Azure client ID from environment variable."""
    client_id = os.getenv("AZURE_CLIENT_ID")
    if not client_id:
        raise Exception("AZURE_CLIENT_ID environment variable not set")
    return client_id

def save_tokens(access_tok: str, refresh_tok: str = None, expires_in: int = 3600) -> None:
    """Save tokens to disk for persistence across sessions."""
    global access_token, refresh_token, token_expires_at
    
    access_token = access_tok
    if refresh_tok:
        refresh_token = refresh_tok
    token_expires_at = time.time() + expires_in - 300  # 5 min buffer
    
    # Only save to disk if caching is enabled
    if not TOKEN_CACHE_ENABLED:
        logger.info("Token caching disabled - tokens will not persist across sessions")
        return
    
    try:
        token_data = {
            "access_token": access_token,
            "refresh_token": refresh_token,
            "expires_at": token_expires_at
        }
        
        with open(TOKEN_CACHE_FILE, 'w') as f:
            json.dump(token_data, f)
        
        # Set secure permissions (user read/write only)
        TOKEN_CACHE_FILE.chmod(0o600)
        logger.info(f"Tokens saved to {TOKEN_CACHE_FILE}")
        
    except Exception as e:
        logger.warning(f"Failed to save tokens: {e}")

def load_tokens() -> bool:
    """Load tokens from disk. Returns True if valid tokens loaded."""
    global access_token, refresh_token, token_expires_at
    
    # Don't load tokens if caching is disabled
    if not TOKEN_CACHE_ENABLED:
        logger.info("Token caching disabled - will not load cached tokens")
        return False
    
    try:
        if not TOKEN_CACHE_FILE.exists():
            logger.info(f"No token cache file found at {TOKEN_CACHE_FILE}")
            return False
            
        with open(TOKEN_CACHE_FILE, 'r') as f:
            token_data = json.load(f)
        
        access_token = token_data.get("access_token")
        refresh_token = token_data.get("refresh_token")
        token_expires_at = token_data.get("expires_at")
        
        # Check if token is still valid
        if token_expires_at and time.time() < token_expires_at:
            logger.info(f"Valid tokens loaded from {TOKEN_CACHE_FILE}")
            return True
        else:
            logger.info("Cached tokens expired")
            return False
            
    except Exception as e:
        logger.warning(f"Failed to load tokens: {e}")
        return False

async def refresh_access_token() -> bool:
    """Try to refresh the access token using the refresh token."""
    global access_token, msal_app
    
    if not refresh_token or not msal_app:
        return False
    
    try:
        # Try to get accounts from MSAL cache
        accounts = msal_app.get_accounts()
        
        if accounts:
            # Try silent acquisition first
            result = msal_app.acquire_token_silent(SCOPES, account=accounts[0])
            
            if result and "access_token" in result:
                save_tokens(
                    result["access_token"],
                    result.get("refresh_token", refresh_token),
                    result.get("expires_in", 3600)
                )
                logger.info("Token refreshed successfully")
                return True
        
        logger.info("Token refresh failed - need new authentication")
        return False
        
    except Exception as e:
        logger.warning(f"Token refresh error: {e}")
        return False

def init_msal_app(client_id: str) -> PublicClientApplication:
    """Initialize MSAL application for authentication."""
    # Create a simple in-memory cache for MSAL
    return PublicClientApplication(
        client_id=client_id,
        authority="https://login.microsoftonline.com/common"
    )

async def ensure_valid_token() -> bool:
    """Ensure we have a valid access token, refreshing if needed."""
    global access_token, msal_app
    
    # First, try loading cached tokens
    if not access_token:
        load_tokens()
    
    # Check if current token is still valid
    if access_token and token_expires_at and time.time() < token_expires_at:
        return True
    
    # Try to refresh the token
    if not msal_app:
        msal_app = init_msal_app(get_client_id())
    
    if await refresh_access_token():
        return True
    
    # No valid token available
    access_token = None
    return False

# Global variable to store the current authentication flow
current_flow = None

@mcp.tool()
async def start_authentication() -> str:
    """
    Start the full authentication process.
    
    Returns:
        Authentication instructions with device code
    """
    global access_token, msal_app, current_flow
    
    try:
        client_id = get_client_id()
        logger.info(f"Starting authentication with client_id: {client_id[:8]}...")
        
        # Create MSAL app if not exists
        if not msal_app:
            msal_app = init_msal_app(client_id)
        
        # Start device code flow
        logger.info("Initiating device flow for authentication...")
        flow = msal_app.initiate_device_flow(scopes=SCOPES)
        
        if "user_code" not in flow:
            error_msg = flow.get('error_description', 'Unknown error in device flow')
            raise Exception(f"Failed to create device flow: {error_msg}")
        
        # Return the authentication instructions
        result = {
            "status": "authentication_required",
            "instructions": f"Go to {flow['verification_uri']} and enter code: {flow['user_code']}",
            "verification_uri": flow['verification_uri'],
            "user_code": flow['user_code'],
            "expires_in": flow.get('expires_in', 900),
            "message": "Please complete authentication, then call 'complete_authentication'"
        }
        
        # Store the flow for completion
        current_flow = flow
        
        return json.dumps(result, indent=2)
        
    except Exception as e:
        logger.error(f"Start authentication error: {str(e)}")
        return json.dumps({
            "status": "error",
            "error": str(e)
        }, indent=2)

@mcp.tool()
async def complete_authentication() -> str:
    """
    Complete the authentication process after user enters device code.
    
    Returns:
        Authentication status and user info
    """
    global access_token, msal_app, current_flow
    
    try:
        if not current_flow:
            return json.dumps({
                "status": "error",
                "error": "No authentication flow in progress. Call 'start_authentication' first."
            }, indent=2)
        
        if not msal_app:
            return json.dumps({
                "status": "error", 
                "error": "MSAL app not initialized"
            }, indent=2)
        
        logger.info("Completing device flow authentication...")
        
        # Complete the flow
        result = msal_app.acquire_token_by_device_flow(current_flow)
        
        if "access_token" in result:
            # Save tokens for future use
            save_tokens(
                result["access_token"],
                result.get("refresh_token"),
                result.get("expires_in", 3600)
            )
            
            logger.info("Authentication successful and tokens cached!")
            
            # Test the token with a basic Graph API call
            try:
                user_info = await make_graph_request("/me")
                return json.dumps({
                    "status": "success",
                    "message": "Authentication completed successfully and tokens cached for future use",
                    "user": user_info.get("displayName", "Unknown"),
                    "email": user_info.get("mail") or user_info.get("userPrincipalName", "Unknown")
                }, indent=2)
                        
            except Exception as graph_error:
                return json.dumps({
                    "status": "partial_success",
                    "message": "Got access token but Graph API test failed",
                    "graph_error": str(graph_error)
                }, indent=2)
        else:
            error_desc = result.get('error_description', 'Unknown authentication error')
            return json.dumps({
                "status": "error",
                "error": f"Authentication failed: {error_desc}"
            }, indent=2)
            
    except Exception as e:
        logger.error(f"Complete authentication error: {str(e)}")
        return json.dumps({
            "status": "error",
            "error": str(e)
        }, indent=2)
    finally:
        # Clear the flow
        current_flow = None

@mcp.tool()
async def check_authentication() -> str:
    """
    Check current authentication status and token validity.
    
    Returns:
        Authentication status information
    """
    try:
        cache_status = "enabled" if TOKEN_CACHE_ENABLED else "disabled"
        cache_file_exists = TOKEN_CACHE_FILE.exists() if TOKEN_CACHE_ENABLED else False
        
        if await ensure_valid_token():
            try:
                user_info = await make_graph_request("/me")
                time_until_expiry = int(token_expires_at - time.time()) if token_expires_at else 0
                
                return json.dumps({
                    "status": "authenticated",
                    "user": user_info.get("displayName", "Unknown"),
                    "email": user_info.get("mail") or user_info.get("userPrincipalName", "Unknown"),
                    "token_valid_for_seconds": max(0, time_until_expiry),
                    "token_valid_for_hours": round(max(0, time_until_expiry) / 3600, 1),
                    "token_caching": cache_status,
                    "cache_file_exists": cache_file_exists,
                    "cache_file_path": str(TOKEN_CACHE_FILE) if TOKEN_CACHE_ENABLED else "N/A"
                }, indent=2)
                
            except Exception as graph_error:
                return json.dumps({
                    "status": "token_invalid",
                    "error": str(graph_error),
                    "message": "Token exists but API call failed - may need re-authentication",
                    "token_caching": cache_status
                }, indent=2)
        else:
            return json.dumps({
                "status": "not_authenticated",
                "message": "No valid authentication token. Please call 'start_authentication'",
                "token_caching": cache_status,
                "cache_file_exists": cache_file_exists
            }, indent=2)
            
    except Exception as e:
        return json.dumps({
            "status": "error",
            "error": str(e),
            "token_caching": "unknown"
        }, indent=2)

async def make_graph_request(endpoint: str, method: str = "GET", data: Dict = None) -> Dict:
    """Make a request to Microsoft Graph API."""
    # Ensure we have a valid token before making the request
    if not await ensure_valid_token():
        raise Exception("Not authenticated. Please call 'start_authentication' and 'complete_authentication' first.")
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    url = f"{GRAPH_BASE_URL}{endpoint}"
    
    async with httpx.AsyncClient() as client:
        if method == "GET":
            response = await client.get(url, headers=headers)
        elif method == "POST":
            response = await client.post(url, headers=headers, json=data)
        elif method == "PATCH":
            response = await client.patch(url, headers=headers, json=data)
        else:
            raise ValueError(f"Unsupported HTTP method: {method}")
    
    if response.status_code >= 400:
        raise Exception(f"Graph API error: {response.status_code} - {response.text}")
    
    return response.json()

@mcp.tool()
async def list_notebooks() -> str:
    """
    List all OneNote notebooks.
    
    Returns:
        JSON string containing notebook information
    """
    try:
        logger.info("Making request to /me/onenote/notebooks")
        notebooks = await make_graph_request("/me/onenote/notebooks")
        logger.info(f"Graph API response received with {len(notebooks.get('value', []))} notebooks")
        
        result = []
        for notebook in notebooks.get("value", []):
            result.append({
                "id": notebook.get("id"),
                "name": notebook.get("displayName"),
                "created": notebook.get("createdDateTime"),
                "modified": notebook.get("lastModifiedDateTime")
            })
        
        logger.info(f"Returning {len(result)} notebooks")
        return json.dumps(result, indent=2)
    
    except Exception as e:
        logger.error(f"Error in list_notebooks: {str(e)}")
        return f"Error listing notebooks: {str(e)}"

@mcp.tool()
async def list_sections(notebook_id: str) -> str:
    """
    List sections in a specific notebook.
    
    Args:
        notebook_id: ID of the notebook to list sections from
    
    Returns:
        JSON string containing section information
    """
    try:
        sections = await make_graph_request(f"/me/onenote/notebooks/{notebook_id}/sections")
        
        result = []
        for section in sections.get("value", []):
            result.append({
                "id": section.get("id"),
                "name": section.get("displayName"),
                "created": section.get("createdDateTime"),
                "modified": section.get("lastModifiedDateTime")
            })
        
        return json.dumps(result, indent=2)
    
    except Exception as e:
        return f"Error listing sections: {str(e)}"

@mcp.tool()
async def list_pages(section_id: str) -> str:
    """
    List pages in a specific section.
    
    Args:
        section_id: ID of the section to list pages from
    
    Returns:
        JSON string containing page information
    """
    try:
        pages = await make_graph_request(f"/me/onenote/sections/{section_id}/pages")
        
        result = []
        for page in pages.get("value", []):
            result.append({
                "id": page.get("id"),
                "title": page.get("title"),
                "created": page.get("createdDateTime"),
                "modified": page.get("lastModifiedDateTime"),
                "content_url": page.get("contentUrl")
            })
        
        return json.dumps(result, indent=2)
    
    except Exception as e:
        return f"Error listing pages: {str(e)}"

@mcp.tool()
async def get_page_content(page_id: str) -> str:
    """
    Get the content of a specific page.
    
    Args:
        page_id: ID of the page to retrieve content from
    
    Returns:
        Page content as HTML or error message
    """
    try:
        # Get page content (returns HTML)
        async with httpx.AsyncClient() as client:
            headers = {"Authorization": f"Bearer {access_token}"}
            response = await client.get(
                f"{GRAPH_BASE_URL}/me/onenote/pages/{page_id}/content",
                headers=headers
            )
            
            if response.status_code >= 400:
                return f"Error getting page content: {response.status_code} - {response.text}"
            
            return response.text
    
    except Exception as e:
        return f"Error getting page content: {str(e)}"

@mcp.tool()
async def clear_token_cache() -> str:
    """
    Clear the stored authentication tokens.
    
    Returns:
        Status message
    """
    global access_token, refresh_token, token_expires_at
    
    try:
        # Clear in-memory tokens
        access_token = None
        refresh_token = None
        token_expires_at = None
        
        # Remove cache file
        if TOKEN_CACHE_FILE.exists():
            TOKEN_CACHE_FILE.unlink()
            
        return json.dumps({
            "status": "success",
            "message": "Token cache cleared. You will need to re-authenticate."
        }, indent=2)
        
    except Exception as e:
        return json.dumps({
            "status": "error",
            "error": str(e)
        }, indent=2)

def main():
    """Main entry point for the server."""
    # Log token caching configuration
    cache_status = "enabled" if TOKEN_CACHE_ENABLED else "disabled"
    logger.info(f"OneNote MCP Server starting - Token caching: {cache_status}")
    
    if TOKEN_CACHE_ENABLED:
        logger.info(f"Token cache file: {TOKEN_CACHE_FILE}")
        # Try to load cached tokens on startup
        if load_tokens():
            logger.info("Cached tokens loaded successfully")
        else:
            logger.info("No valid cached tokens found")
    
    mcp.run()

if __name__ == "__main__":
    main()
