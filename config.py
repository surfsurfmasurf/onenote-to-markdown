"""
config.py
---------
Loads environment variables from the .env file and exposes them
as module-level constants used throughout the application.
"""

import os
from dotenv import load_dotenv

# Load variables from .env file (if it exists)
load_dotenv()

# --- Microsoft Identity Platform ---
# Application (client) ID from Azure App Registration
CLIENT_ID: str = os.getenv("CLIENT_ID", "")

# Client secret generated in Azure App Registration
CLIENT_SECRET: str = os.getenv("CLIENT_SECRET", "")

# "consumers" for personal Microsoft accounts (MSA),
# or your Azure AD tenant ID for work/school accounts.
TENANT_ID: str = os.getenv("TENANT_ID", "consumers")

# --- Output ---
# Root directory where the exported Markdown files will be written.
OUTPUT_DIR: str = os.getenv("OUTPUT_DIR", "./output")

# --- Microsoft Graph API ---
# OAuth 2.0 scopes required to read OneNote data.
SCOPES: list[str] = ["Notes.Read", "Notes.Read.All", "offline_access"]

# Base URL for all Microsoft Graph API v1.0 endpoints.
GRAPH_API_BASE: str = "https://graph.microsoft.com/v1.0"
