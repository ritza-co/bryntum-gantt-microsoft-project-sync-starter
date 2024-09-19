import {
    PublicClientApplication,
    InteractionRequiredAuthError
} from '@azure/msal-browser';

const msalConfig = {
    auth : {
        clientId  : import.meta.env.VITE_MICROSOFT_ENTRA_APP_ID,
        authority : `https://login.microsoftonline.com/${
      import.meta.env.VITE_MICROSOFT_ENTRA_TENANT_ID
    }`,
        redirectUri : 'http://localhost:5173'
    }
};

// Initialize a PublicClientApplication object.
const msalInstance =
  await PublicClientApplication.createPublicClientApplication(msalConfig);

const msalRequest = { scopes : [`https://${
    import.meta.env.VITE_MICROSOFT_DYNAMICS_ORG_ID
  }.api.crm4.dynamics.com/.default`] };

export function ensureScope(scope) {
    if (
        !msalRequest.scopes.some((s) => s.toLowerCase() === scope.toLowerCase())
    ) {
        msalRequest.scopes.push(scope);
    }
}


// Log the user in
export async function signIn() {
    const authResult = await msalInstance.loginPopup(msalRequest);
    sessionStorage.setItem('msalAccount', authResult.account.username);
}

export async function getToken() {
    const account = sessionStorage.getItem('msalAccount');
    if (!account) {
        throw new Error(
            'User info cleared from session. Please sign out and sign in again.'
        );
    }
    try {
    // First, attempt to get the token silently
        const silentRequest = {
            scopes  : msalRequest.scopes,
            account : msalInstance.getAccountByUsername(account)
        };

        const silentResult = await msalInstance.acquireTokenSilent(silentRequest);
        return silentResult.accessToken;
    }
    catch (silentError) {
    // If silent request fails with InteractionRequiredAuthError,
    // attempt to get the token interactively
        if (silentError instanceof InteractionRequiredAuthError) {
            const interactiveResult = await msalInstance.acquireTokenPopup(
                msalRequest
            );
            return interactiveResult.accessToken;
        }
        else {
            throw silentError;
        }
    }
}
