import { PublicClientApplication, InteractionRequiredAuthError } from '@azure/msal-browser';
import { CONFIG } from './config.js';

// Thrown when a token for extra scopes cannot be acquired silently — the
// user (or an admin) has to consent to the additional permission first.
export class ConsentRequiredError extends Error {
    constructor(scopes, cause) {
        super('Additional permissions are required: ' + scopes.map(s => s.split('/').pop()).join(', '));
        this.name = 'ConsentRequiredError';
        this.scopes = scopes;
        this.cause = cause;
    }
}

class AuthService {
    constructor() {
        this.msalInstance = null;
        this.account = null;
        this.isInitialized = false;
    }

    async initialize() {
        if (this.isInitialized) return;

        try {
            this.msalInstance = new PublicClientApplication({
                auth: {
                    clientId: CONFIG.msal.clientId,
                    authority: CONFIG.msal.authority,
                    redirectUri: CONFIG.msal.redirectUri
                },
                cache: {
                    cacheLocation: 'sessionStorage',
                    storeAuthStateInCookie: false
                }
            });

            await this.msalInstance.initialize();
            this.isInitialized = true;
            
            // Handle redirect response
            // In MSAL v5, handleRedirectPromise throws no_token_request_cache_error
            // on normal page loads (no redirect in progress) — treat this as null response.
            let response = null;
            try {
                response = await this.msalInstance.handleRedirectPromise();
            } catch (redirectError) {
                if (redirectError?.errorCode !== 'no_token_request_cache_error') {
                    throw redirectError;
                }
            }

            if (response && response.account) {
                this.account = response.account;
                console.log('Authentication successful after redirect');
            } else {
                // Try to get account from cache
                const accounts = this.msalInstance.getAllAccounts();
                if (accounts.length > 0) {
                    this.account = accounts[0];
                    this.msalInstance.setActiveAccount(this.account);
                    console.log('Account found in cache');
                }
            }
        } catch (error) {
            console.error('Error initializing MSAL:', error);
            throw new Error('Failed to initialize authentication');
        }
    }

    async login() {
        if (!this.isInitialized) {
            throw new Error('Auth service not initialized');
        }

        try {
            // Use redirect flow instead of popup to avoid popup issues
            await this.msalInstance.loginRedirect({
                scopes: CONFIG.msal.scopes
            });
        } catch (error) {
            console.error('Login failed:', error);
            throw new Error('Login failed: ' + error.message);
        }
    }

    async logout() {
        if (!this.isInitialized) return;

        try {
            await this.msalInstance.logoutRedirect({
                account: this.account
            });
            this.account = null;
        } catch (error) {
            console.error('Logout failed:', error);
            throw new Error('Logout failed: ' + error.message);
        }
    }

    async getAccessToken() {
        if (!this.account) {
            throw new Error('No account available');
        }

        try {
            const response = await this.msalInstance.acquireTokenSilent({
                scopes: CONFIG.msal.scopes,
                account: this.account
            });
            
            return response.accessToken;
        } catch (error) {
            // Try interactive token acquisition using redirect
            try {
                await this.msalInstance.acquireTokenRedirect({
                    scopes: CONFIG.msal.scopes,
                    account: this.account
                });
                // After redirect, this will be handled by handleRedirectPromise
                return null;
            } catch (interactiveError) {
                console.error('Token acquisition failed:', interactiveError);
                throw new Error('Failed to acquire access token');
            }
        }
    }

    // Acquire a token for specific scopes without ever redirecting. Used for
    // incremental consent: callers show their own "grant access" UI when this
    // throws ConsentRequiredError instead of the page navigating away.
    async getAccessTokenForScopes(scopes) {
        if (!this.account) {
            throw new Error('No account available');
        }

        try {
            const response = await this.msalInstance.acquireTokenSilent({
                scopes,
                account: this.account
            });
            return response.accessToken;
        } catch (error) {
            const consentCodes = ['consent_required', 'interaction_required', 'invalid_grant', 'login_required'];
            if (error instanceof InteractionRequiredAuthError || consentCodes.includes(error?.errorCode)) {
                throw new ConsentRequiredError(scopes, error);
            }
            throw error;
        }
    }

    // Redirect to Entra ID to consent to additional scopes. The page reloads
    // afterwards, so callers should persist any state they want to resume.
    async requestAdditionalScopes(scopes) {
        if (!this.isInitialized) {
            throw new Error('Auth service not initialized');
        }
        await this.msalInstance.acquireTokenRedirect({
            scopes,
            account: this.account
        });
    }

    isAuthenticated() {
        return this.account !== null;
    }

    getAccount() {
        return this.account;
    }

    getTenantId() {
        if (!this.account) {
            return null;
        }
        return this.account.tenantId;
    }

    getTenantInfo() {
        if (!this.account) {
            return null;
        }
        
        return {
            tenantId: this.account.tenantId,
            homeAccountId: this.account.homeAccountId,
            username: this.account.username,
            name: this.account.name,
            environment: this.account.environment
        };
    }
}

// Create singleton instance
export const authService = new AuthService();
export default authService;
