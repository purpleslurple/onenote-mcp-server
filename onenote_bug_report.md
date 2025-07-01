# Bug Report: OneNote MCP Server Returns Unusable Page URLs

## Summary
The OneNote MCP server returns Graph API content URLs instead of user-accessible OneNote page URLs, making cross-page linking unusable from within OneNote.

## Problem Description
When creating or retrieving pages, the MCP server returns URLs in this format:
```
https://graph.microsoft.com/v1.0/users/matsch@sasites.com/onenote/pages/0-cadfb407915b42c9ab7aa73af7ba8341!248-F4873BF54ECA028!s6f5705bb804c4cf680d42e575dd7768c/content
```

These URLs are Graph API endpoints, not user-facing links. They don't work when:
- Clicked from within OneNote
- Shared with other users
- Used for cross-page references

## Expected Behavior
The MCP server should return OneNote web app URLs that users can actually navigate to, such as:
```
https://www.onenote.com/notebooks/{notebook-id}/sections/{section-id}/pages/{page-id}
```

## Impact
- **Cross-page linking broken**: Cannot create functional references between pages
- **Poor user experience**: Links appear to work but lead to API errors
- **Knowledge management limitation**: Defeats the purpose of connected note-taking

## Steps to Reproduce
1. Create a page using `onenote:create_page`
2. Note the returned `content_url` in the response
3. Try to navigate to that URL from OneNote web app or desktop
4. Observe that it doesn't work as a user-facing link

## Technical Details
- **Current**: Returns Graph API `/content` endpoints
- **Needed**: User-navigable OneNote web URLs
- **Root cause**: MCP server using internal API URLs instead of public page URLs

## Affected Functions
- `onenote:create_page` - Returns unusable `content_url`
- `onenote:list_pages` - Returns unusable `content_url` for each page
- Any cross-page linking functionality

## Suggested Fix
The Graph API response likely contains both the content URL (for API access) and a web URL (for user navigation). The MCP server should:
1. Extract the user-facing web URL from Graph API responses
2. Return that URL in addition to or instead of the content URL
3. Provide a separate field for API content URLs if needed for internal operations

## Priority
**Medium-High** - This limits the core knowledge management use case of creating connected, navigable note networks.

## Workaround
Currently none - users cannot create functional cross-page links through the MCP server.

## Environment
- OneNote MCP Server version: [current]
- Platform: macOS
- OneNote: Web app
- Graph API: v1.0