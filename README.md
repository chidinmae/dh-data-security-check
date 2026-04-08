# M365 Security Audit Tool & AI Governance PRD

## Overview
This tool is a proof-of-concept for an M365 Security Audit dashboard. It is designed to demonstrate technical proficiency in integrating Microsoft Graph API with modern React/TypeScript architectures, specifically focusing on data security and compliance within the Microsoft 365 ecosystem.

The primary goal is to identify over-privileged sharing links in SharePoint and scan for sensitive information (specifically Dutch BSN numbers) to mitigate data leakage risks, which is a critical component of AI Governance (ensuring AI models don't ingest or expose sensitive data).

## Key Features
- **Entra ID Authentication**: Secure login via MSAL.
- **SharePoint Site Discovery**: Automatically lists all sites within the tenant.
- **Selective Scanning**: Default "Scan All" with options to target specific sites.
- **Permission Analysis**: Identifies files shared via "Anonymous" or "Everyone" links.
- **PII Scanning**: Scans text-based file contents for Dutch BSN (Citizen Service Numbers).
- **Compliance Dashboard**: Premium UI for real-time risk assessment and reporting.
- **Exportable Reports**: Results can be exported to JSON/CSV for further analysis.

## AI Governance Context
In the age of Copilot and LLMs, data security is paramount. This tool addresses the "Oversharing" problem—where AI agents might surface sensitive information because user permissions were set too broadly. By auditing and remediating these links, organizations can safely implement AI solutions.

## Technical Stack
- **Frontend**: React 19, Vite, TypeScript.
- **Styling**: Vanilla CSS (Modern design system).
- **Auth**: `@azure/msal-browser`, `@azure/msal-react`.
- **API**: Microsoft Graph SDK.
- **Icons**: Lucide React.

## Configuration
To run this tool, you must register an application in the [Azure Portal](https://portal.azure.com/) with the following configuration:

### Redirect URI
- `http://localhost:5173/`

### API Permissions (Application or Delegated)
- `User.Read`
- `Sites.Read.All`
- `Files.Read.All`

### Environment Variables
Create a `.env` file in the root directory:
```env
VITE_CLIENT_ID=your-client-id
VITE_TENANT_ID=your-tenant-id
```

## Getting Started
1. `npm install`
2. Configure `.env`
3. `npm run dev`

---
*Created for technical demonstration purposes.*
