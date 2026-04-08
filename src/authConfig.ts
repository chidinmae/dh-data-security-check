import type { Configuration, PopupRequest } from "@azure/msal-browser";

// Configuration object for MSAL
export const msalConfig: Configuration = {
    auth: {
        clientId: import.meta.env.VITE_CLIENT_ID || "00000000-0000-0000-0000-000000000000",
        authority: `https://login.microsoftonline.com/${import.meta.env.VITE_TENANT_ID || "common"}`,
        redirectUri: "http://localhost:5173",
        postLogoutRedirectUri: "/",
    },
    cache: {
        cacheLocation: "sessionStorage",
    }
};

// Scopes for API calls
export const loginRequest: PopupRequest = {
    scopes: ["User.Read", "Sites.Read.All", "Files.Read.All"]
};

export const graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
    graphSitesEndpoint: "https://graph.microsoft.com/v1.0/sites?search="
};
