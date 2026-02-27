# onenote-to-markdown

> Migrate all your Microsoft OneNote notebooks to **Obsidian** or **Notion** with a single command.

Connects to your OneNote account via the Microsoft Graph API, downloads every notebook / section / page, converts the HTML content to clean Markdown files, and organises them in a folder structure that mirrors your original OneNote hierarchy.

## Features

- Authenticates via **Device Code Flow** - works in any terminal, no browser redirect required on the server
- Exports **all notebooks, sections, and pages** automatically
- Generates **YAML front matter** (title, created, modified, source) compatible with Obsidian and Notion
- Creates an **_index.md table of contents** with Obsidian wikilinks per notebook
- **Handles pagination** - works correctly even with hundreds of pages
- **Duplicate-safe** - appends a numeric suffix when two pages share the same title
- **Token caching** - stores the OAuth token locally so you do not have to log in on every run

## Prerequisites

- Python 3.10 or higher
- A Microsoft account with OneNote notes
- An Azure App Registration (free)

## Setup

### 1. Register an Azure Application

1. Go to the Azure Portal (https://portal.azure.com) and sign in.
2. Navigate to Azure Active Directory > App registrations > New registration.
3. Set Redirect URI to Public client/native: https://login.microsoftonline.com/common/oauth2/nativeclient
4. Go to API permissions and add: Notes.Read, Notes.Read.All, offline_access
5. Grant admin consent.
6. Go to Certificates & secrets > New client secret and copy the Value.

### 2. Configure the tool

Copy .env.example to .env and fill in your credentials:

CLIENT_ID=your-application-client-id
CLIENT_SECRET=your-client-secret-value
TENANT_ID=consumers
OUTPUT_DIR=./output

### 3. Install dependencies

pip install -r requirements.txt

## Usage

python main.py

On the first run you will see a device login prompt. Open the URL shown, enter the code, and sign in with your Microsoft account.

## Importing into Obsidian

Open Obsidian > Open folder as vault > select the output/ directory.

## Importing into Notion

Settings > Import > Markdown & CSV > select the output/ directory.

## Project structure

| File | Purpose |
|---|---|
| main.py | Entry point |
| onenote_fetcher.py | Microsoft Graph API client + authentication |
| converter.py | HTML to Markdown conversion |
| exporter.py | Writes .md files to disk |
| config.py | Loads settings from .env |
| requirements.txt | Python dependencies |
| .env.example | Template for environment variables |

## License

MIT
