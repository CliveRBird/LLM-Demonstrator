# Microsoft 365 Authentication Web App

A complete web application that demonstrates Microsoft 365 work/school account authentication using MSAL.js (Microsoft Authentication Library for JavaScript).

## Features

- ✅ Microsoft 365 work/school account authentication
- ✅ User profile display with photo
- ✅ Access token acquisition
- ✅ Microsoft Graph API integration
- ✅ Responsive design
- ✅ Error handling and fallback authentication methods
- ✅ Both popup and redirect authentication flows

## Setup Instructions

### 1. Azure AD App Registration

1. Go to [Azure AD App Registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Click **"New registration"**
3. Configure the application:
   - **Name**: "Microsoft 365 Auth Demo" (or your preferred name)
   - **Supported account types**: "Accounts in any organizational directory" 
   - **Redirect URI**: 
     - Platform: "Single-page application (SPA)"
     - URI: `http://localhost:3000` (for local development) or your domain
4. Click **"Register"**
5. Copy the **Application (client) ID** from the overview page

### 2. Configure API Permissions (Optional)

For additional Microsoft Graph functionality:

1. Go to **API permissions** in your app registration
2. Click **"Add a permission"**
3. Select **Microsoft Graph** → **Delegated permissions**
4. Add permissions like:
   - `User.Read` (basic profile)
   - `Mail.Read` (read user's mail)
   - `Calendars.Read` (read user's calendar)
5. Click **"Grant admin consent"** if required by your organization

### 3. Update Configuration

1. Open `app.js`
2. Replace `"YOUR_CLIENT_ID_HERE"` with your actual client ID:
   ```javascript
   const msalConfig = {
       auth: {
           clientId: "your-actual-client-id-here",
           // ... rest of config
       }
   };
   ```

### 4. Run the Application

#### Option A: Local Web Server
```bash
# Using Python (if installed)
python -m http.server 3000

# Using Node.js (if installed)
npx http-server -p 3000

# Using PHP (if installed)
php -S localhost:3000
```

#### Option B: Live Server (VS Code Extension)
1. Install the "Live Server" extension in VS Code
2. Right-click on `index.html`
3. Select "Open with Live Server"

#### Option C: Direct File Access
Open `index.html` directly in your browser (may have limitations with some authentication flows)

## How It Works

### Authentication Flow

1. **Sign In**: Uses MSAL.js to authenticate with Microsoft 365 accounts
2. **Token Management**: Automatically handles token acquisition and renewal
3. **Profile Loading**: Retrieves user information and photo from Microsoft Graph
4. **API Calls**: Demonstrates calling Microsoft Graph APIs with proper authentication

### Supported Account Types

This app supports:
- ✅ Microsoft 365 work accounts (user@company.com)
- ✅ School accounts (user@school.edu)
- ❌ Personal Microsoft accounts (user@outlook.com) - use Quick Authentication for these

### Key Components

- **MSAL Configuration**: Configured for organizational accounts
- **Authentication Methods**: Popup-first with redirect fallback
- **Token Handling**: Silent token renewal with interactive fallback
- **Error Handling**: Comprehensive error handling and user feedback
- **Microsoft Graph Integration**: Profile data and API call examples

## File Structure

```
microsoft-365-auth-app/
├── index.html          # Main HTML page
├── styles.css          # Styling and responsive design
├── app.js              # MSAL.js authentication logic
└── README.md           # This file
```

## Troubleshooting

### Common Issues

1. **"Application not found" error**
   - Verify the client ID is correct
   - Ensure the app registration exists in the correct Azure AD tenant

2. **Redirect URI mismatch**
   - Verify the redirect URI in Azure AD matches your app's URL exactly
   - Include both `http://localhost:3000` and your production domain

3. **Popup blocked**
   - The app automatically falls back to redirect authentication
   - Allow popups for better user experience

4. **CORS errors**
   - Serve the app from a web server, not file:// protocol
   - Ensure your domain is properly configured in Azure AD

5. **Permission errors**
   - Check that required API permissions are granted
   - Some permissions may require admin consent

### Browser Compatibility

- ✅ Chrome 90+
- ✅ Firefox 88+
- ✅ Safari 14+
- ✅ Edge 90+
- ⚠️ Internet Explorer (limited support, requires polyfills)

## Security Considerations

- Client ID is safe to expose in client-side code
- Never include client secrets in client-side applications
- Tokens are stored securely in browser storage
- HTTPS recommended for production deployments
- Regular token refresh ensures security

## Customization

### Styling
- Modify `styles.css` to match your brand colors and design
- The app uses Microsoft's Fluent Design principles

### Functionality
- Add more Microsoft Graph API calls in `app.js`
- Customize the required scopes in the configuration
- Add additional user interface elements as needed

### Deployment
- Update redirect URIs for production domains
- Consider using Azure Static Web Apps or similar services
- Implement proper error logging and monitoring

## Resources

- [MSAL.js Documentation](https://docs.microsoft.com/en-us/azure/active-directory/develop/msal-overview)
- [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/)
- [Azure AD App Registration](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Microsoft Identity Platform](https://docs.microsoft.com/en-us/azure/active-directory/develop/)# LLM-Demonstrator
