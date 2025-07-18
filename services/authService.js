const { ConfidentialClientApplication } = require('@azure/msal-node');
require('dotenv').config();

class AuthService {
    constructor() {
        this.msalInstance = new ConfidentialClientApplication({
            auth: {
                clientId: process.env.CLIENT_ID,
                clientSecret: process.env.CLIENT_SECRET,
                authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`
            }
        });
    }

    async getAccessToken() {
        try {
            const clientCredentialRequest = {
                scopes: ['https://graph.microsoft.com/.default']
            };

            const response = await this.msalInstance.acquireTokenByClientCredential(clientCredentialRequest);
           
            return response.accessToken;
        } catch (error) {
            console.error('Error getting access token:', error);
            throw error;
        }
    }
}

module.exports = new AuthService();