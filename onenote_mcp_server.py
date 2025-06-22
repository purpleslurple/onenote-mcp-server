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

# Global variables for authentication
access_token: Optional[str] = None
msal_app: Optional[PublicClientApplication] = None

def get_client_id() -> str:
    """Get the Azure client ID from environment variable."""
    client_id = os.getenv("AZURE_CLIENT_ID")
    if not client_id:
        raise Exception("AZURE_CLIENT_ID environment variable not set")
    return client_id

def init_msal_app(client_id: str) -> PublicClientApplication:
    """Initialize MSAL application for authentication."""
    return PublicClientApplication(
        client_id=client_id,
        authority="https://login.microsoftonline.com/common"
    )

async def get_access_token(client_id: str) -> str:
    """Get access token using device code flow."""
    global access_token, msal_app
    
    if access_token:
        return access_token
    
    try:
        if not msal_app:
            msal_app = init_msal_app(client_id)
        
        # Start device code flow
        flow = msal_app.initiate_device_flow(scopes=SCOPES)
        
        if "user_code" not in flow:
            error_msg = flow.get('error_description', 'Unknown error in device flow')
            raise Exception(f"Failed to create device flow: {error_msg}")
        
        logger.info(f"To authenticate, go to: {flow['verification_uri']}")
        logger.info(f"And enter code: {flow['user_code']}")
        
        # Complete the flow
        result = msal_app.acquire_token_by_device_flow(flow)
        
        if "access_token" in result:
            access_token = result["access_token"]
            logger.info("Authentication successful!")
            return access_token
        else:
            error_desc = result.get('error_description', 'Unknown authentication error')
            raise Exception(f"Authentication failed: {error_desc}")
            
    except Exception as e:
        logger.error(f"Authentication error: {str(e)}")
        raise

async def make_graph_request(endpoint: str, method: str = "GET", data: Dict = None) -> Dict:
    """Make a request to Microsoft Graph API."""
    if not access_token:
        raise Exception("Not authenticated. Please call get_access_token first.")
    
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
            access_token = result["access_token"]
            logger.info("Authentication successful!")
            
            # Test the token with a basic Graph API call
            try:
                user_info = await make_graph_request("/me")
                return json.dumps({
                    "status": "success",
                    "message": "Authentication completed successfully",
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
async def list_notebooks() -> str:
    """
    List all OneNote notebooks.
    
    Returns:
        JSON string containing notebook information
    """
    try:
        if not access_token:
            return json.dumps({
                "status": "error",
                "error": "Not authenticated. Please call 'start_authentication' and 'complete_authentication' first."
            }, indent=2)
        
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
        if not access_token:
            return json.dumps({
                "status": "error",
                "error": "Not authenticated. Please call 'start_authentication' and 'complete_authentication' first."
            }, indent=2)
            
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
        if not access_token:
            return json.dumps({
                "status": "error",
                "error": "Not authenticated. Please call 'start_authentication' and 'complete_authentication' first."
            }, indent=2)
            
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
async def test_authentication() -> str:
    """
    Test OneNote authentication and basic connectivity.
    
    Returns:
        Authentication status and basic user info
    """
    try:
        client_id = get_client_id()
        logger.info(f"Testing authentication with client_id: {client_id[:8]}...")
        
        # Test getting access token
        token = await get_access_token(client_id)
        logger.info("Access token obtained successfully")
        
        # Test basic Graph API call
        user_info = await make_graph_request("/me")
        logger.info("User info retrieved successfully")
        
        return json.dumps({
            "status": "success",
            "user": user_info.get("displayName", "Unknown"),
            "email": user_info.get("mail") or user_info.get("userPrincipalName", "Unknown"),
            "message": "Authentication working correctly"
        }, indent=2)
        
    except Exception as e:
        logger.error(f"Authentication test failed: {str(e)}")
        return json.dumps({
            "status": "error",
            "error": str(e),
            "message": "Authentication failed"
        }, indent=2)

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
        if not access_token:
            return json.dumps({
                "status": "error",
                "error": "Not authenticated. Please call 'start_authentication' and 'complete_authentication' first."
            }, indent=2)
        
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

def main():
    """Main entry point for the server."""
    mcp.run()

if __name__ == "__main__":
    main()
