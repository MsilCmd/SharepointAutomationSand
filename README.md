# SharepointAutomationSand
A Sharepoint automation sandbox
# SharePoint Automation Suite

A comprehensive Python toolkit for SharePoint automation using Microsoft Graph API, Office365-REST-Python-Client, and Dropbox SDK.

## Features

- **Authentication** — Azure AD OAuth2 (client credentials & delegated)
- **SharePoint Lists** — Full CRUD on list items
- **Document Management** — Upload, download, and manage files
- **Content Migration** — Bulk migrate content between sites/libraries
- **Reporting Dashboards** — Generate HTML/Excel reports on usage and data
- **Provisioning** — Automate site, list, and permission provisioning
- **Dropbox → SharePoint** — Automated import pipeline

## Project Structure

```
sharepoint-automation/
├── src/
│   ├── auth/               # Azure AD + Dropbox authentication
│   ├── sharepoint/         # Graph API + REST client wrappers
│   ├── dropbox/            # Dropbox SDK integration
│   ├── reporting/          # Dashboard and report generation
│   ├── provisioning/       # Site/list provisioning automation
│   └── migration/          # Content migration utilities
├── tests/                  # Unit and integration tests
├── scripts/                # CLI entry-point scripts
├── config/                 # Configuration templates
└── docs/                   # Additional documentation
```

## Prerequisites

- Python 3.10+
- An Azure AD App Registration with the following Graph API permissions:
  - `Sites.ReadWrite.All`
  - `Files.ReadWrite.All`
  - `Lists.ReadWrite.All`
  - `User.Read.All` (for reporting)
- A Dropbox App with `files.content.read` scope

## Quick Start

### 1. Clone & Install

```bash
git clone https://github.com/your-org/sharepoint-automation.git
cd sharepoint-automation
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Configure

```bash
cp config/.env.example .env
# Fill in your credentials in .env
```

### 3. Run Examples

```bash
# Read a SharePoint list
python scripts/list_items.py --site "your-site" --list "Tasks"

# Upload a document
python scripts/upload_doc.py --site "your-site" --lib "Documents" --file "./report.pdf"

# Run the Dropbox → SharePoint import
python scripts/dropbox_import.py --dropbox-path "/Reports" --sp-lib "Imported"

# Generate a usage report
python scripts/generate_report.py --output ./dashboard.html
```

## Environment Variables

| Variable | Description |
|---|---|
| `AZURE_TENANT_ID` | Your Azure AD tenant ID |
| `AZURE_CLIENT_ID` | App registration client ID |
| `AZURE_CLIENT_SECRET` | App registration client secret |
| `SHAREPOINT_SITE_URL` | Base SharePoint site URL (e.g. `https://contoso.sharepoint.com/sites/mysite`) |
| `DROPBOX_ACCESS_TOKEN` | Dropbox OAuth2 access token |
| `DROPBOX_APP_KEY` | Dropbox app key (for refresh token flow) |
| `DROPBOX_APP_SECRET` | Dropbox app secret |
| `DROPBOX_REFRESH_TOKEN` | Dropbox refresh token |

## Architecture

```
┌─────────────────────────────────┐
│         CLI / Scripts           │
└────────────┬────────────────────┘
             │
┌────────────▼────────────────────┐
│        Core Services            │
│  ┌──────────┐  ┌─────────────┐  │
│  │SharePoint│  │   Dropbox   │  │
│  │ Manager  │  │   Manager   │  │
│  └────┬─────┘  └──────┬──────┘  │
│       │               │         │
│  ┌────▼───────────────▼──────┐  │
│  │      Auth Manager         │  │
│  │  (Azure AD + Dropbox)     │  │
│  └───────────────────────────┘  │
└─────────────────────────────────┘
             │
┌────────────▼────────────────────┐
│     External APIs               │
│  Microsoft Graph  |  Dropbox    │
│  SP REST API      |  SDK        │
└─────────────────────────────────┘
```

## License

MIT
