# SharePoint Express POC

This is a Node.js Express application that connects to SharePoint using Azure AD authentication. It provides APIs to:
- List files and folders from SharePoint
- Download files to local storage
- Upload files to Azure Blob Storage
- Track downloaded files across storage locations

## API Documentation

### 1. Health Check
Verify if the application is running.

```bash
curl http://localhost:3000/
```

### 2. List Files (Root)
Get all files from the root of the SharePoint site.

```bash
curl http://localhost:3000/files
```

### 3. Download Specific File
Download a specific file by its name.

```bash
curl http://localhost:3000/download/example.pdf > example.pdf
```

### 4. List Folder Contents
List all files and folders in a specific path. The path parameter is optional.

```bash
# List root contents
curl http://localhost:3000/list

# List specific folder
curl http://localhost:3000/list/Documents

# List nested folder
curl http://localhost:3000/list/Documents/Reports
```

Response format:
```json
{
  "currentPath": "Documents/Reports",
  "items": [
    {
      "name": "example.pdf",
      "type": "file",
      "path": "Documents/Reports/example.pdf",
      "lastModified": "2023-01-01T00:00:00Z",
      "size": 1234567,
      "webUrl": "https://sharepoint.com/path/to/file",
      "mimeType": "application/pdf",
      "downloadUrl": "https://...."
    },
    {
      "name": "SubFolder",
      "type": "folder",
      "path": "Documents/Reports/SubFolder",
      "lastModified": "2023-01-01T00:00:00Z",
      "size": 0,
      "webUrl": "https://sharepoint.com/path/to/folder",
      "childCount": 5
    }
  ]
}
```

### 5. Download Folder Contents
Download all files from a folder recursively to either local storage or Azure Blob Storage.

#### a. Download to Local Storage
```bash
curl -X POST http://localhost:3000/download-folder \
  -H "Content-Type: application/json" \
  -d '{
    "folderPath": "Documents/Reports",
    "destination": "local"
  }'
```

#### b. Upload to Azure Blob Storage
```bash
curl -X POST http://localhost:3000/download-folder \
  -H "Content-Type: application/json" \
  -d '{
    "folderPath": "Documents/Reports",
    "destination": "blob"
  }'
```

Response format:
```json
{
  "status": "success",
  "totalFiles": 10,
  "files": [
    {
      "file": "Documents/Reports/example.pdf",
      "status": "success",
      "destination": "/local/path/or/blob/url"
    }
  ]
}
```

## Prerequisites

1. Node.js installed on your machine
2. Azure AD App Registration with appropriate permissions:
   - Microsoft Graph API permissions:
     - Files.Read.All
     - Sites.Read.All
3. SharePoint site URL
4. Azure AD credentials:
   - Tenant ID
   - Client ID
   - Client Secret
5. Azure Blob Storage (optional, for blob storage support):
   - Storage Account Connection String
   - Container Name

## Setup Instructions

1. Clone this repository
2. Install dependencies:
   ```bash
   npm install
   ```

3. Configure environment variables:
   - Copy `.env.example` to `.env`
   - Update the following variables in `.env`:
     ```env
     # Azure AD Configuration
     AZURE_TENANT_ID=your_tenant_id
     AZURE_CLIENT_ID=your_client_id
     AZURE_CLIENT_SECRET=your_client_secret
     SHAREPOINT_SITE_URL=your_sharepoint_site_url
     PORT=3000

     # Azure Blob Storage Configuration (Optional)
     AZURE_STORAGE_CONNECTION_STRING=your_storage_connection_string
     AZURE_STORAGE_CONTAINER_NAME=your_container_name

     # Local Download Path
     LOCAL_DOWNLOAD_PATH=./downloads
     ```

4. Azure AD App Configuration:
   - Go to Azure Portal > Azure Active Directory > App Registrations
   - Register a new application or use an existing one
   - Add the following API permissions:
     - Microsoft Graph
       - Files.Read.All
       - Sites.Read.All
   - Generate a client secret and save it securely

5. Azure Blob Storage Configuration (Optional):
   - Create an Azure Storage Account if you don't have one
   - Create a container for storing SharePoint files
   - Get the connection string from Access Keys section
   - Update the `.env` file with the connection string and container name

6. Start the application:
   ```bash
   # Development mode with auto-reload
   npm run dev

   # Production mode
   npm start
   ```

## API Documentation

### 1. Health Check
Verify if the application is running.

```bash
curl http://localhost:3000/
```

### 2. List Files (Root)
Get all files from the root of the SharePoint site.

```bash
curl http://localhost:3000/files
```

### 3. Download Specific File
Download a specific file by its name.

```bash
curl http://localhost:3000/download/example.pdf > example.pdf
```

### 4. List Folder Contents
List all files and folders in a specific path. The path parameter is optional.

```bash
# List root contents from all document libraries
curl http://localhost:3000/list

# List specific folder
curl http://localhost:3000/list/Documents

# List nested folder
curl http://localhost:3000/list/Documents/Reports
```

Response format:
```json
{
  "currentPath": "Documents/Reports",
  "drives": [
    {
      "id": "drive1Id",
      "name": "Documents",
      "webUrl": "https://sharepoint.com/..."
    }
  ],
  "items": [
    {
      "name": "example.pdf",
      "type": "file",
      "path": "Documents/Reports/example.pdf",
      "lastModified": "2023-01-01T00:00:00Z",
      "size": 1234567,
      "webUrl": "https://sharepoint.com/...",
      "driveId": "drive1Id",
      "driveName": "Documents",
      "mimeType": "application/pdf",
      "downloadUrl": "https://...",
      "isFile": true
    },
    {
      "name": "SubFolder",
      "type": "folder",
      "path": "Documents/Reports/SubFolder",
      "lastModified": "2023-01-01T00:00:00Z",
      "size": 0,
      "webUrl": "https://sharepoint.com/...",
      "driveId": "drive1Id",
      "driveName": "Documents",
      "childCount": 5,
      "isFolder": true
    }
  ]
}
```

### 5. Download Folder Contents
Download all files from a folder recursively to either local storage or Azure Blob Storage.

```bash
# Download to local storage
curl -X POST http://localhost:3000/download-folder \
  -H "Content-Type: application/json" \
  -d '{
    "folderPath": "Documents/Reports",
    "destination": "local"
  }'

# Upload to Azure Blob Storage
curl -X POST http://localhost:3000/download-folder \
  -H "Content-Type: application/json" \
  -d '{
    "folderPath": "Documents/Reports",
    "destination": "blob"
  }'
```

Response format:
```json
{
  "status": "success",
  "totalFiles": 10,
  "files": [
    {
      "file": "Documents/Reports/example.pdf",
      "status": "success",
      "destination": "/local/path/or/blob/url"
    }
  ]
}
```

### 6. Fetch Specific Files
Fetch files from a SharePoint folder with advanced filtering options and upload to local storage or Azure Blob Storage.

```bash
curl -X POST http://localhost:3000/fetch-files \
  -H "Content-Type: application/json" \
  -d '{
    "sourcePath": "Documents/Reports",
    "destination": "local",
    "filePatterns": ["*.pdf", "doc*.docx"],
    "recursive": true,
    "dateRange": {
      "from": "2023-01-01",
      "to": "2023-12-31"
    },
    "preserveFolderStructure": true
  }'
```

Request parameters:
- `sourcePath`: SharePoint folder path (optional, defaults to root)
- `destination`: Either 'local' or 'blob' (optional, defaults to 'local')
- `filePatterns`: Array of file patterns using wildcards (optional)
- `recursive`: Whether to search in subfolders (optional, defaults to false)
- `dateRange`: Filter files by modification date (optional)
  - `from`: Start date in YYYY-MM-DD format
  - `to`: End date in YYYY-MM-DD format
- `preserveFolderStructure`: Maintain folder structure in destination (optional, defaults to true)

Response format:
```json
{
  "status": "success",
  "totalFiles": 5,
  "sourcePath": "Documents/Reports",
  "destination": "local",
  "filePatterns": ["*.pdf", "doc*.docx"],
  "recursive": true,
  "dateRange": {
    "from": "2023-01-01",
    "to": "2023-12-31"
  },
  "preserveFolderStructure": true,
  "files": [
    {
      "file": "Documents/Reports/example.pdf",
      "status": "success",
      "destination": "/local/path/or/blob/url",
      "size": 1234567,
      "lastModified": "2023-01-01T00:00:00Z"
    }
  ]
}
```

### 7. List Downloaded Files
List all files that have been downloaded to local storage or uploaded to blob storage.

```bash
# List all downloaded files
curl http://localhost:3000/downloads

# List only local files
curl http://localhost:3000/downloads?type=local

# List only blob storage files
curl http://localhost:3000/downloads?type=blob
```

Response format:
```json
{
  "totalFiles": 10,
  "storageTypes": ["local", "blob"],
  "files": {
    "local": [
      {
        "name": "example.pdf",
        "path": "Documents/Reports/example.pdf",
        "size": 1234567,
        "lastModified": "2023-01-01T00:00:00Z",
        "type": "local",
        "fullPath": "/absolute/path/to/file"
      }
    ],
    "blob": [
      {
        "name": "example.pdf",
        "path": "Documents/Reports/example.pdf",
        "size": 1234567,
        "lastModified": "2023-01-01T00:00:00Z",
        "type": "blob",
        "url": "https://storage.blob.core.windows.net/...",
        "contentType": "application/pdf"
      }
    ]
  }
}
```

## Error Handling

All endpoints return appropriate HTTP status codes:
- 200: Success
- 400: Bad Request (invalid parameters)
- 404: Not Found (file or folder not found)
- 500: Internal Server Error

Error responses include a message explaining the error:
```json
{
  "error": "Detailed error message"
}
```

## Security Notes

1. Never commit the `.env` file to version control
2. Rotate Azure AD client secrets periodically
3. Use appropriate CORS policies in production
4. Implement rate limiting for production use
5. Keep Azure AD and Storage access tokens secure
6. Monitor and audit file access patterns

## Available Endpoints

1. GET `/`: Health check endpoint
2. GET `/files`: Lists all files in the root Documents library
3. GET `/download/:filename`: Downloads a specific file by name

## Error Handling

The application includes basic error handling for:
- Authentication failures
- File not found
- Server errors

## Security Notes

- Never commit the `.env` file to version control
- Rotate client secrets periodically
- Use appropriate CORS policies in production
- Implement rate limiting for production use
