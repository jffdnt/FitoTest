import { Msal2Provider } from '@microsoft/mgt-msal2-provider';

interface TokenResponse {
    accessToken: string;
    expiresOn: Date;
    scopes: string[];
    error?: any;
}

export default class MSALWrapper {
    private provider: Msal2Provider;

    constructor(clientID: string, authority: string) {
        console.log('Initializing MSALWrapper with:', { clientID, authority });
        this.provider = new Msal2Provider({
            clientId: clientID,
            authority: authority
        });
    }

    public async handleLoggedInUser(scopes: string[], userEmail: string): Promise<TokenResponse | null> {
        console.log('Attempting to get token for logged in user:', { scopes, userEmail });
        try {
            const accessToken = await this.provider.getAccessToken({ scopes });
            console.log('Token response received:', {
                hasAccessToken: !!accessToken,
                length: accessToken?.length
            });
            
            if (!accessToken) {
                console.warn('Token response missing access token');
                return null;
            }

            return {
                accessToken,
                expiresOn: new Date(Date.now() + 3600000), // Default 1 hour expiration
                scopes
            };
        } catch (error) {
            console.error('Error getting token for logged in user:', {
                error: error instanceof Error ? error.message : error,
                stack: error instanceof Error ? error.stack : undefined
            });
            return null;
        }
    }

    public async acquireAccessToken(scopes: string[], userEmail: string): Promise<TokenResponse | null> {
        console.log('Attempting to acquire new access token:', { scopes, userEmail });
        try {
            const accessToken = await this.provider.getAccessToken({ scopes });
            console.log('New token response received:', {
                hasAccessToken: !!accessToken,
                length: accessToken?.length
            });

            if (!accessToken) {
                console.warn('New token response missing access token');
                return null;
            }

            return {
                accessToken,
                expiresOn: new Date(Date.now() + 3600000), // Default 1 hour expiration
                scopes
            };
        } catch (error) {
            console.error('Error acquiring access token:', {
                error: error instanceof Error ? error.message : error,
                stack: error instanceof Error ? error.stack : undefined
            });
            return null;
        }
    }
}
