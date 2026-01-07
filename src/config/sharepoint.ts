// SharePoint Embedded Configuration
// Update these values with your actual Azure AD and SharePoint Embedded settings

export const SHAREPOINT_CONFIG = {
  CLIENT_ID: "50cbacb0-e16f-4f63-a678-01359bfac87b",
  TENANT_ID: "fc14a141-120b-4368-b125-571da82b7865",
  CONTAINER_TYPE_ID: "9162b1be-e7db-4b0d-bc1a-331df4dea97e",
} as const;

// MSAL Configuration
export const MSAL_CONFIG = {
  auth: {
    clientId: SHAREPOINT_CONFIG.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${SHAREPOINT_CONFIG.TENANT_ID}`,
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
};

// Graph API endpoint
export const GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0";
