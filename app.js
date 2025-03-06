require('isomorphic-fetch');
require('dotenv').config();
const express = require('express');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const app = express();
const port = process.env.PORT || 3000;

// Initialize Azure AD credentials
const credential = new ClientSecretCredential(
    process.env.AZURE_TENANT_ID,
    process.env.AZURE_CLIENT_ID,
    process.env.AZURE_CLIENT_SECRET
);

// Create authentication provider
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
});

// Initialize Graph client
const graphClient = Client.initWithMiddleware({ authProvider });

app.get('/', (req, res) => {
    res.send('SharePoint Express POC is running!');
});

// Example endpoint to list files from a SharePoint site
app.get('/files', async (req, res) => {
    try {
        const siteUrl = new URL(process.env.SHAREPOINT_SITE_URL);
        const hostname = siteUrl.hostname;
        const sitePath = siteUrl.pathname;

        // Get the SharePoint site ID
        const site = await graphClient
            .api(`/sites/${hostname}:${sitePath}`)
            .get();

        // Get files from the Documents library
        const files = await graphClient
            .api(`/sites/${site.id}/drive/root/children`)
            .get();

        res.json(files);
    } catch (error) {
        console.error('Error fetching files:', error);
        res.status(500).json({ error: error.message });
    }
});

// Example endpoint to download a specific file by name
app.get('/download/:filename', async (req, res) => {
    try {
        const siteUrl = new URL(process.env.SHAREPOINT_SITE_URL);
        const hostname = siteUrl.hostname;
        const sitePath = siteUrl.pathname;
        const filename = req.params.filename;

        const site = await graphClient
            .api(`/sites/${hostname}:${sitePath}`)
            .get();

        const items = await graphClient
            .api(`/sites/${site.id}/drive/root/search(q='${filename}')`)
            .get();

        if (items.value.length === 0) {
            return res.status(404).json({ error: 'File not found' });
        }

        const file = items.value[0];
        const download = await graphClient
            .api(`/sites/${site.id}/drive/items/${file.id}/content`)
            .get();

        res.setHeader('Content-Disposition', `attachment; filename=${filename}`);
        res.setHeader('Content-Type', file.file.mimeType);
        download.pipe(res);
    } catch (error) {
        console.error('Error downloading file:', error);
        res.status(500).json({ error: error.message });
    }
});

app.listen(port, () => {
    console.log(`Server is running on port ${port}`);
});
