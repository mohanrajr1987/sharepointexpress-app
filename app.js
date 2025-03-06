require('isomorphic-fetch');
require('dotenv').config();
const fs = require('fs-extra');
const path = require('path');
const { BlobServiceClient } = require('@azure/storage-blob');
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

// Helper function to get all drives in the site
async function getAllDrives(site) {
    try {
        const drives = await graphClient
            .api(`/sites/${site.id}/drives`)
            .get();
        return drives.value;
    } catch (error) {
        console.error('Error getting drives:', error);
        return [];
    }
}

// Helper function to get items from a specific drive and path
async function getItemsFromDrive(driveId, folderPath) {
    try {
        const apiPath = folderPath
            ? `/drives/${driveId}/root:/${folderPath}:/children`
            : `/drives/${driveId}/root/children`;

        const items = await graphClient
            .api(apiPath)
            .expand('folder,file')
            .get();

        return items.value;
    } catch (error) {
        console.error(`Error getting items from drive ${driveId}:`, error);
        return [];
    }
}

// Endpoint to list all folders and files recursively
app.get('/list/*', async (req, res) => {
    try {
        const siteUrl = new URL(process.env.SHAREPOINT_SITE_URL);
        const hostname = siteUrl.hostname;
        const sitePath = siteUrl.pathname;

        // Get the SharePoint site ID
        const site = await graphClient
            .api(`/sites/${hostname}:${sitePath}`)
            .get();

        // Get all drives in the site
        const drives = await getAllDrives(site);
        
        // Get the folder path from the URL
        const folderPath = req.params[0] || '';
        
        let allItems = [];
        
        // Get items from all drives
        for (const drive of drives) {
            const items = await getItemsFromDrive(drive.id, folderPath);
            
            // Process items from this drive
            const processedItems = items.map(item => ({
                name: item.name,
                type: item.folder ? 'folder' : 'file',
                path: folderPath ? `${folderPath}/${item.name}` : item.name,
                lastModified: item.lastModifiedDateTime,
                size: item.size,
                webUrl: item.webUrl,
                driveId: drive.id,
                driveName: drive.name,
                ...(item.folder && { 
                    childCount: item.folder.childCount,
                    isFolder: true
                }),
                ...(item.file && { 
                    mimeType: item.file.mimeType,
                    downloadUrl: item['@microsoft.graph.downloadUrl'],
                    isFile: true
                })
            }));
            
            allItems = allItems.concat(processedItems);
        }

        // Sort items: folders first, then files, both alphabetically
        allItems.sort((a, b) => {
            if (a.type === b.type) {
                return a.name.localeCompare(b.name);
            }
            return a.type === 'folder' ? -1 : 1;
        });

        res.json({
            currentPath: folderPath || 'root',
            drives: drives.map(drive => ({
                id: drive.id,
                name: drive.name,
                webUrl: drive.webUrl
            })),
            items: allItems
        });
    } catch (error) {
        console.error('Error listing items:', error);
        res.status(500).json({ error: error.message });
    }
});

// Helper function to download file to local path
async function downloadToLocal(url, localPath) {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`Failed to download file: ${response.statusText}`);
    await fs.ensureDir(path.dirname(localPath));
    const buffer = await response.buffer();
    await fs.writeFile(localPath, buffer);
    return localPath;
}

// Helper function to upload to blob storage
async function uploadToBlob(blobClient, url) {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`Failed to download file: ${response.statusText}`);
    const buffer = await response.buffer();
    await blobClient.uploadData(buffer);
    return blobClient.url;
}

// Endpoint to download folder contents
app.post('/download-folder', async (req, res) => {
    try {
        const { folderPath = '', destination = 'local' } = req.body;
        const siteUrl = new URL(process.env.SHAREPOINT_SITE_URL);
        const hostname = siteUrl.hostname;
        const sitePath = siteUrl.pathname;

        // Get the SharePoint site ID
        const site = await graphClient
            .api(`/sites/${hostname}:${sitePath}`)
            .get();

        // Function to recursively get all items
        async function getAllItems(currentPath) {
            const apiPath = currentPath
                ? `/sites/${site.id}/drive/root:/${currentPath}:/children`
                : `/sites/${site.id}/drive/root/children`;

            const items = await graphClient
                .api(apiPath)
                .expand('folder,file')
                .get();

            let allItems = [];

            for (const item of items.value) {
                if (item.folder) {
                    const subPath = currentPath ? `${currentPath}/${item.name}` : item.name;
                    const subItems = await getAllItems(subPath);
                    allItems = allItems.concat(subItems);
                } else if (item.file) {
                    allItems.push({
                        name: item.name,
                        path: currentPath ? `${currentPath}/${item.name}` : item.name,
                        downloadUrl: item['@microsoft.graph.downloadUrl']
                    });
                }
            }

            return allItems;
        }

        const allFiles = await getAllItems(folderPath);
        const results = [];

        if (destination === 'blob') {
            // Upload to Azure Blob Storage
            const blobServiceClient = BlobServiceClient.fromConnectionString(
                process.env.AZURE_STORAGE_CONNECTION_STRING
            );
            const containerClient = blobServiceClient.getContainerClient(
                process.env.AZURE_STORAGE_CONTAINER_NAME
            );

            for (const file of allFiles) {
                const blobClient = containerClient.getBlockBlobClient(file.path);
                const blobUrl = await uploadToBlob(blobClient, file.downloadUrl);
                results.push({
                    file: file.path,
                    status: 'success',
                    destination: blobUrl
                });
            }
        } else {
            // Download to local path
            const basePath = process.env.LOCAL_DOWNLOAD_PATH;
            await fs.ensureDir(basePath);

            for (const file of allFiles) {
                const localPath = path.join(basePath, file.path);
                await downloadToLocal(file.downloadUrl, localPath);
                results.push({
                    file: file.path,
                    status: 'success',
                    destination: localPath
                });
            }
        }

        res.json({
            status: 'success',
            totalFiles: results.length,
            files: results
        });
    } catch (error) {
        console.error('Error downloading folder:', error);
        res.status(500).json({ error: error.message });
    }
});

app.listen(port, () => {
    console.log(`Server is running on port ${port}`);
});
