// SharePoint Embedded Configuration
// Update these values with your actual Azure AD and SharePoint Embedded settings

export const SHAREPOINT_CONFIG = {
  CLIENT_ID: "[uuid]",
  TENANT_ID: "[uuid]",
  CONTAINER_TYPE_ID: "[uuid]",
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
