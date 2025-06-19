# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a client-side Microsoft 365 authentication web application demonstrating MSAL.js (Microsoft Authentication Library for JavaScript) integration. The app provides a complete authentication flow for Microsoft 365 work/school accounts with Microsoft Graph API integration.

**Key Technologies:**
- MSAL.js 2.38.1 (Microsoft Authentication Library)
- Vanilla JavaScript (no frameworks)
- Microsoft Graph API
- Pure HTML/CSS/JS implementation

## Development Commands

This is a static web application with no build process. Development requires serving the files through a web server:

```bash
# Using Python (most common)
python3 -m http.server 3000

# Using Node.js
npx http-server -p 3000

# Using PHP
php -S localhost:3000
```

**Important:** The app must be served from `http://localhost:3000` to match the Azure AD redirect URI configuration.

## Architecture

### Authentication Flow
The app implements a popup-first authentication strategy with redirect fallback:

1. **Primary Flow:** `msalInstance.loginPopup()` in `app.js:92`
2. **Fallback Flow:** `msalInstance.loginRedirect()` in `app.js:100` when popup fails
3. **Token Management:** Silent token refresh with interactive fallback in `app.js:149-160`

### Core Components

**MSAL Configuration** (`app.js:4-14`):
- Multi-tenant authority (`/common` endpoint)
- Session storage for token caching
- Dynamic redirect URI based on current origin

**Authentication States** (`app.js:264-283`):
- Signed out state: Shows sign-in button
- Signed in state: Shows user profile and action buttons
- Error state: Displays authentication or API errors

**Microsoft Graph Integration** (`app.js:167-197`):
- User profile retrieval (`/me` endpoint)
- User photo loading (`/me/photo/$value` endpoint)
- Access token display and API response formatting

### File Structure
```
├── index.html          # Main UI with authentication sections
├── app.js              # MSAL.js authentication logic and Graph API calls
├── styles.css          # Microsoft Fluent Design-inspired styling
└── README.md           # Setup instructions and troubleshooting
```

## Configuration Requirements

**Critical Setup Step:** Replace `"YOUR_CLIENT_ID_HERE"` in `app.js:6` with actual Azure AD application client ID.

**Azure AD App Registration Requirements:**
- Application type: Single-page application (SPA)
- Redirect URI: `http://localhost:3000` (development) or production domain
- Supported accounts: "Accounts in any organizational directory"
- Required permissions: `User.Read` (minimum)

## Key Implementation Details

**Error Handling Strategy:**
- Popup blocked → Automatic redirect fallback
- Silent token acquisition failure → Interactive popup fallback
- Graph API errors → User-friendly error display
- Multiple authentication attempts → Interaction state management prevents overlapping requests
- Nested popup context → Automatic redirect usage when in popup/iframe

**Security Features:**
- Tokens stored in sessionStorage (configurable to localStorage)
- No client secrets (SPA pattern)
- Automatic token refresh handling
- HTTPS enforcement for production

**Responsive Design:**
- Mobile-first CSS with breakpoints at 768px
- Flexible button groups and user profile layout
- Microsoft Fluent Design color scheme and typography

## Common Development Tasks

**Adding New Graph API Calls:**
1. Add required scopes to `tokenRequest.scopes` in `app.js:22`
2. Implement API call following pattern in `callMicrosoftGraph()` function
3. Add corresponding UI elements in `index.html`

**Modifying Authentication Scopes:**
- Update `loginRequest.scopes` for sign-in permissions
- Update `tokenRequest.scopes` for API access permissions
- Ensure Azure AD app registration includes required permissions

**Styling Customization:**
- Microsoft brand colors defined in `styles.css:11,69,129`
- Fluent Design components follow established patterns
- CSS custom properties can be added for theme consistency

## Debugging Support

Global debugging object available in browser console:
```javascript
window.msalApp.getCurrentAccount()  // Get current user account
window.msalApp.getAccessToken()     // Get current access token
```

**Common Issues:**
- "Application not found" → Verify client ID in `app.js:6`
- CORS errors → Ensure proper web server (not file:// protocol)
- Popup blocked → App automatically falls back to redirect flow
- Token refresh issues → Check browser console for MSAL errors