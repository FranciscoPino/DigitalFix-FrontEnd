import { LogLevel } from '@azure/msal-browser'
import type { Configuration } from '@azure/msal-browser'

const tenantId = import.meta.env.VITE_AZURE_TENANT_ID || 'common'
const clientId = import.meta.env.VITE_AZURE_CLIENT_ID || 'replace-with-client-id'
const apiClientId = import.meta.env.VITE_AZURE_API_CLIENT_ID || clientId

export const msalConfig: Configuration = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: import.meta.env.VITE_REDIRECT_URI || window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: 'sessionStorage',
  },
  system: {
    loggerOptions: {
      loggerCallback: (_level, message, containsPii) => {
        if (!containsPii && import.meta.env.DEV) console.debug(message)
      },
      logLevel: LogLevel.Warning,
      piiLoggingEnabled: false,
    },
  },
}

export const loginRequest = {
  scopes: [import.meta.env.VITE_API_SCOPE || `api://${apiClientId}/access_as_user`],
}
