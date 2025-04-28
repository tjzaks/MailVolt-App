const { google } = require('googleapis');
const { OAuth2Client } = require('google-auth-library');

class GSuiteClient {
    constructor(clientId, clientSecret, redirectUri) {
        this.oauth2Client = new OAuth2Client(
            clientId,
            clientSecret,
            redirectUri
        );
        this.gmail = null;
    }

    async authenticate(authCode) {
        try {
            const { tokens } = await this.oauth2Client.getToken(authCode);
            this.oauth2Client.setCredentials(tokens);
            this.gmail = google.gmail({ version: 'v1', auth: this.oauth2Client });
            return true;
        } catch (error) {
            console.error('G-Suite authentication failed:', error);
            return false;
        }
    }

    async validateEmailDomain(email) {
        try {
            const domain = email.split('@')[1];
            const response = await this.gmail.users.settings.sendAs.list({
                userId: 'me'
            });
            
            const allowedDomains = response.data.sendAs.map(sendAs => 
                sendAs.sendAsEmail.split('@')[1]
            );
            
            return allowedDomains.includes(domain);
        } catch (error) {
            console.error('Email domain validation failed:', error);
            return false;
        }
    }

    async getDistributionLists() {
        try {
            const response = await this.gmail.users.settings.filters.list({
                userId: 'me'
            });
            // Process and return distribution lists
            return response.data.filters || [];
        } catch (error) {
            console.error('Failed to get distribution lists:', error);
            return [];
        }
    }
}

module.exports = GSuiteClient; 