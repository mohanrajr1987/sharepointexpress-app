# SharePoint Express POC

This is a Node.js Express application that connects to SharePoint using Azure AD authentication to fetch and download files.

## Prerequisites

1. Node.js installed on your machine
2. Azure AD App Registration with appropriate permissions
3. SharePoint site URL
4. Azure AD credentials (Tenant ID, Client ID, and Client Secret)

## Setup Instructions

1. Clone this repository
2. Install dependencies:
   ```bash
   npm install
   ```

3. Configure environment variables:
   - Copy `.env.example` to `.env`
   - Update the following variables in `.env`:
     - AZURE_TENANT_ID: Your Azure AD tenant ID
     - AZURE_CLIENT_ID: Your Azure AD application (client) ID
     - AZURE_CLIENT_SECRET: Your Azure AD client secret
     - SHAREPOINT_SITE_URL: Your SharePoint site URL
     - PORT: Port number for the Express server (default: 3000)

4. Azure AD App Configuration:
   - Go to Azure Portal > Azure Active Directory > App Registrations
   - Register a new application or use an existing one
   - Add the following API permissions:
     - Microsoft Graph
       - Files.Read.All
       - Sites.Read.All
   - Generate a client secret and save it securely

5. Start the application:
   ```bash
   # Development mode with auto-reload
   npm run dev

   # Production mode
   npm start
   ```

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
