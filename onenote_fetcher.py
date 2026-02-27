"""
onenote_fetcher.py
------------------
Handles Microsoft authentication (Device Code Flow) and all
Microsoft Graph API calls needed to retrieve OneNote content:
notebooks, sections, pages, and raw page HTML.
"""

import json
import requests
import msal

from config import CLIENT_ID, TENANT_ID, SCOPES, GRAPH_API_BASE

# Path for MSAL's persistent token cache so users don't have to
# re-authenticate on every run.
_TOKEN_CACHE_PATH = "msal_token_cache.bin"


def _load_token_cache() -> msal.SerializableTokenCache:
    """Load the MSAL token cache from disk (if it exists)."""
    cache = msal.SerializableTokenCache()
    try:
        with open(_TOKEN_CACHE_PATH, "r") as f:
            cache.deserialize(f.read())
    except FileNotFoundError:
        pass  # First run - no cache yet
    return cache


def _save_token_cache(cache: msal.SerializableTokenCache) -> None:
    """Persist the MSAL token cache to disk after any state change."""
    if cache.has_state_changed:
        with open(_TOKEN_CACHE_PATH, "w") as f:
            f.write(cache.serialize())


def get_access_token() -> str:
    """
    Obtain a valid Microsoft Graph access token.

    Uses the Device Code Flow so the tool can run in any terminal
    (no browser redirect needed on the machine running the script).
    Tokens are cached locally so subsequent runs skip re-authentication.

    Returns:
        A bearer token string ready to be used in Authorization headers.

    Raises:
        RuntimeError: If authentication fails for any reason.
    """
    cache = _load_token_cache()

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )

    # Try silent token acquisition from cache first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_token_cache(cache)
            print("Authenticated using cached token.")
            return result["access_token"]

    # Fall back to Device Code Flow (interactive)
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(
            "Failed to start device flow: " + json.dumps(flow, indent=2)
        )

    print("\n" + "=" * 60)
    print("  Microsoft Login Required")
    print(f"  1. Open: {flow['verification_uri']}")
    print(f"  2. Enter code: {flow['user_code']}")
    print("=" * 60 + "\n")

    result = app.acquire_token_by_device_flow(flow)
    _save_token_cache(cache)

    if "access_token" not in result:
        raise RuntimeError(
            "Authentication failed: "
            + result.get("error_description", "Unknown error")
        )

    print("Login successful!\n")
    return result["access_token"]


class OneNoteFetcher:
    """
    Wraps Microsoft Graph API calls for OneNote resources.

    Attributes:
        token: The OAuth 2.0 bearer token used in every request.
        headers: Common HTTP headers attached to every API call.
    """

    def __init__(self, access_token: str) -> None:
        self.token = access_token
        self.headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json",
        }

    def _get_paged(self, url: str, params: dict | None = None) -> list:
        """
        Perform a GET request and automatically follow @odata.nextLink
        pagination so callers always receive the full result set.
        """
        items: list = []
        while url:
            response = requests.get(url, headers=self.headers, params=params)
            response.raise_for_status()
            data = response.json()
            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")  # None when last page reached
            params = None  # nextLink already contains query params
        return items

    def get_notebooks(self) -> list[dict]:
        """Return all OneNote notebooks accessible by the signed-in user."""
        print("Fetching notebooks...")
        return self._get_paged(f"{GRAPH_API_BASE}/me/onenote/notebooks")

    def get_sections(self, notebook_id: str) -> list[dict]:
        """Return all sections inside a given notebook."""
        url = f"{GRAPH_API_BASE}/me/onenote/notebooks/{notebook_id}/sections"
        return self._get_paged(url)

    def get_pages(self, section_id: str) -> list[dict]:
        """Return all pages (metadata only) inside a given section."""
        url = f"{GRAPH_API_BASE}/me/onenote/sections/{section_id}/pages"
        return self._get_paged(url)

    def get_page_content(self, page_id: str) -> str:
        """
        Download the raw HTML content of a single OneNote page.

        Returns:
            HTML string representing the page body.
        """
        url = f"{GRAPH_API_BASE}/me/onenote/pages/{page_id}/content"
        response = requests.get(url, headers=self.headers)
        response.raise_for_status()
        return response.text

    def get_all_structure(self) -> list[dict]:
        """
        Walk the entire OneNote hierarchy and return a nested structure
        of notebooks, sections, and page metadata.
        """
        structure: list[dict] = []

        for notebook in self.get_notebooks():
            print(f"  Notebook: {notebook['displayName']}")
            nb_entry: dict = {
                "id": notebook["id"],
                "name": notebook["displayName"],
                "sections": [],
            }

            for section in self.get_sections(notebook["id"]):
                print(f"    Section: {section['displayName']}")
                sec_entry: dict = {
                    "id": section["id"],
                    "name": section["displayName"],
                    "pages": [],
                }

                for page in self.get_pages(section["id"]):
                    sec_entry["pages"].append(
                        {
                            "id": page["id"],
                            "title": page.get("title") or "Untitled",
                            "createdDateTime": page.get("createdDateTime"),
                            "lastModifiedDateTime": page.get(
                                "lastModifiedDateTime"
                            ),
                        }
                    )

                nb_entry["sections"].append(sec_entry)

            structure.append(nb_entry)

        return structure
