// Microsoft 365 Authentication App using MSAL.js

// MSAL Configuration
const msalConfig = {
    auth: {
        clientId: "68a386f4-a324-4b94-b89d-f74fda77a6fc", // Replace with your Azure AD app client ID
        authority: "https://login.microsoftonline.com/f7d8edcf-0316-4e4a-822f-944bd9eeafaf", // Multi-tenant endpoint
        redirectUri: window.location.origin, // Current page URL
    },
    cache: {
        cacheLocation: "sessionStorage", // Can be "localStorage" or "sessionStorage"
        storeAuthStateInCookie: false, // Set to true for IE 11 or Edge compatibility
    }
};

// Scopes for Microsoft Graph API
const loginRequest = {
    scopes: ["openid", "profile", "email", "User.Read"]
};

const tokenRequest = {
    scopes: ["User.Read", "Mail.Read"]
};

// Initialize MSAL instance
const msalInstance = new msal.PublicClientApplication(msalConfig);

// DOM elements
const signInBtn = document.getElementById('sign-in-btn');
const signOutBtn = document.getElementById('sign-out-btn');
const signOutInitialBtn = document.getElementById('sign-out-initial-btn');
const getTokenBtn = document.getElementById('get-token-btn');
const callGraphBtn = document.getElementById('call-graph-btn');
const signInSection = document.getElementById('sign-in-section');
const signedInSection = document.getElementById('signed-in-section');
const tokenSection = document.getElementById('token-section');
const graphSection = document.getElementById('graph-section');
const errorSection = document.getElementById('error-section');

// User info elements
const userName = document.getElementById('user-name');
const userEmail = document.getElementById('user-email');
const userOrg = document.getElementById('user-org');
const userPhoto = document.getElementById('user-photo');
const accessTokenDisplay = document.getElementById('access-token');
const graphResponse = document.getElementById('graph-response');
const errorInfo = document.getElementById('error-info');

// Global variables
let currentAccount = null;
let accessToken = null;

// Initialize the application
async function initializeApp() {
    try {
        await msalInstance.initialize();
        
        // Handle redirect response if returning from authentication
        const response = await msalInstance.handleRedirectPromise();
        if (response) {
            handleAuthResponse(response);
        } else {
            // Check if user is already signed in
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                currentAccount = accounts[0];
                showSignedInState();
                await loadUserProfile();
            }
        }
    } catch (error) {
        console.error('Failed to initialize MSAL:', error);
        showError('Failed to initialize authentication: ' + error.message);
    }
}

// Handle authentication response
function handleAuthResponse(response) {
    if (response && response.account) {
        currentAccount = response.account;
        msalInstance.setActiveAccount(currentAccount);
        showWelcomeNotification();
        showSignedInState();
        loadUserProfile();
    }
}

// Sign in function
async function signIn() {
    try {
        hideError();
        
        // Check if we need to handle popup or redirect
        const response = await msalInstance.loginPopup(loginRequest);
        handleAuthResponse(response);
    } catch (error) {
        console.error('Sign in failed:', error);
        
        if (error.errorCode === 'popup_window_error' || error.errorCode === 'popup_canceled') {
            // Fallback to redirect if popup fails
            try {
                await msalInstance.loginRedirect(loginRequest);
            } catch (redirectError) {
                console.error('Redirect sign in failed:', redirectError);
                showError('Sign in failed: ' + redirectError.message);
            }
        } else {
            showError('Sign in failed: ' + error.message);
        }
    }
}

// Sign out function
async function signOut() {
    try {
        hideError();
        
        const logoutRequest = {
            account: currentAccount,
            postLogoutRedirectUri: window.location.origin
        };
        
        await msalInstance.logoutPopup(logoutRequest);
        currentAccount = null;
        accessToken = null;
        showSignedOutState();
    } catch (error) {
        console.error('Sign out failed:', error);
        
        // Fallback to redirect logout
        const logoutRequest = {
            account: currentAccount,
            postLogoutRedirectUri: window.location.origin
        };
        
        try {
            await msalInstance.logoutRedirect(logoutRequest);
        } catch (redirectError) {
            console.error('Redirect sign out failed:', redirectError);
            showError('Sign out failed: ' + redirectError.message);
        }
    }
}

// Get access token
async function getAccessToken() {
    try {
        hideError();
        
        const request = {
            ...tokenRequest,
            account: currentAccount
        };
        
        // Try silent token acquisition first
        try {
            const response = await msalInstance.acquireTokenSilent(request);
            accessToken = response.accessToken;
            displayAccessToken(accessToken);
        } catch (silentError) {
            console.warn('Silent token acquisition failed, falling back to popup:', silentError);
            
            // Fall back to popup
            const response = await msalInstance.acquireTokenPopup(request);
            accessToken = response.accessToken;
            displayAccessToken(accessToken);
        }
    } catch (error) {
        console.error('Token acquisition failed:', error);
        showError('Failed to get access token: ' + error.message);
    }
}

// Call Microsoft Graph API
async function callMicrosoftGraph() {
    try {
        hideError();
        
        if (!accessToken) {
            await getAccessToken();
        }
        
        if (!accessToken) {
            throw new Error('No access token available');
        }
        
        const response = await fetch('https://graph.microsoft.com/v1.0/me', {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (!response.ok) {
            throw new Error(`Graph API call failed: ${response.status} ${response.statusText}`);
        }
        
        const graphData = await response.json();
        displayGraphResponse(graphData);
    } catch (error) {
        console.error('Microsoft Graph call failed:', error);
        showError('Microsoft Graph call failed: ' + error.message);
    }
}

// Load user profile information
async function loadUserProfile() {
    try {
        if (currentAccount) {
            // Display basic account info
            userName.textContent = currentAccount.name || 'Unknown User';
            userEmail.textContent = currentAccount.username || 'No email available';
            
            // Try to get organization info from claims
            if (currentAccount.idTokenClaims) {
                const claims = currentAccount.idTokenClaims;
                userOrg.textContent = claims.organization || claims.tid || 'Organization info not available';
            }
            
            // Try to load user photo
            await loadUserPhoto();
        }
    } catch (error) {
        console.error('Failed to load user profile:', error);
    }
}

// Load user photo from Microsoft Graph
async function loadUserPhoto() {
    try {
        // Get token for Graph API call
        const request = {
            scopes: ["User.Read"],
            account: currentAccount
        };
        
        const response = await msalInstance.acquireTokenSilent(request);
        
        // Get user photo
        const photoResponse = await fetch('https://graph.microsoft.com/v1.0/me/photo/$value', {
            headers: {
                'Authorization': `Bearer ${response.accessToken}`
            }
        });
        
        if (photoResponse.ok) {
            const photoBlob = await photoResponse.blob();
            const photoUrl = URL.createObjectURL(photoBlob);
            userPhoto.src = photoUrl;
            userPhoto.style.display = 'block';
        }
    } catch (error) {
        console.log('Could not load user photo:', error.message);
        // Photo is optional, so we don't show this as an error
    }
}

// Display access token
function displayAccessToken(token) {
    accessTokenDisplay.value = token;
    tokenSection.style.display = 'block';
}

// Display Microsoft Graph response
function displayGraphResponse(data) {
    graphResponse.textContent = JSON.stringify(data, null, 2);
    graphSection.style.display = 'block';
}

// Show welcome notification
function showWelcomeNotification() {
    console.log('Showing welcome notification');
    const welcomeNotification = document.getElementById('welcome-notification');
    if (welcomeNotification) {
        welcomeNotification.style.display = 'block';
        console.log('Welcome notification displayed');
        
        // Auto-hide after 5 seconds
        setTimeout(() => {
            hideWelcomeNotification();
        }, 5000);
    } else {
        console.error('Welcome notification element not found');
    }
}

// Hide welcome notification
function hideWelcomeNotification() {
    const welcomeNotification = document.getElementById('welcome-notification');
    welcomeNotification.style.display = 'none';
}

// Show already signed in state (user was already authenticated)
function showAlreadySignedInState() {
    signInBtn.style.display = 'none';
    signOutInitialBtn.style.display = 'inline-block';
    signInSection.style.display = 'block';
    signedInSection.style.display = 'none';
}

// Show signed in state (after fresh authentication)
function showSignedInState() {
    signInSection.style.display = 'none';
    signedInSection.style.display = 'block';
}

// Show signed out state
function showSignedOutState() {
    signInBtn.style.display = 'inline-block';
    signOutInitialBtn.style.display = 'none';
    signInSection.style.display = 'block';
    signedInSection.style.display = 'none';
    tokenSection.style.display = 'none';
    graphSection.style.display = 'none';
    hideWelcomeNotification();
    
    // Clear user info
    userName.textContent = '';
    userEmail.textContent = '';
    userOrg.textContent = '';
    userPhoto.style.display = 'none';
    accessTokenDisplay.value = '';
    graphResponse.textContent = '';
}

// Show error message
function showError(message) {
    errorInfo.textContent = message;
    errorSection.style.display = 'block';
}

// Hide error message
function hideError() {
    errorSection.style.display = 'none';
}

// Toggle setup instructions
function toggleSetup() {
    const setupDetails = document.getElementById('setup-details');
    if (setupDetails.style.display === 'none') {
        setupDetails.style.display = 'block';
    } else {
        setupDetails.style.display = 'none';
    }
}

// Event listeners
signInBtn.addEventListener('click', signIn);
signOutBtn.addEventListener('click', signOut);
signOutInitialBtn.addEventListener('click', signOut);
getTokenBtn.addEventListener('click', getAccessToken);
callGraphBtn.addEventListener('click', callMicrosoftGraph);

// Initialize the app when the page loads
document.addEventListener('DOMContentLoaded', initializeApp);

// Handle page visibility changes to refresh tokens if needed
document.addEventListener('visibilitychange', () => {
    if (!document.hidden && currentAccount) {
        // Page became visible, check if we need to refresh tokens
        console.log('Page became visible, checking authentication state');
    }
});

// Export functions for global access (useful for debugging)
window.msalApp = {
    signIn,
    signOut,
    getAccessToken,
    callMicrosoftGraph,
    toggleSetup,
    getCurrentAccount: () => currentAccount,
    getAccessToken: () => accessToken
};