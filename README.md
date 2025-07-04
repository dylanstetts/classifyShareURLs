# Classify Graph API Share URLs

This Python script helps developers and support engineers inspect and validate Microsoft Graph sharing links.

## Features

- Automatically acquires an app-only access token using MSAL.
- Encodes and decodes sharing URLs to and from Graph-compatible `shareId` format.
- Classifies shared resources (e.g., document, list item) based on URL patterns.
- Calls the appropriate Microsoft Graph API endpoint (`/driveItem` or `/listItem`) based on classification.
- Handles throttling (HTTP 429) and transient errors with retry logic.
- Prints detailed metadata and item responses for inspection.

## Requirements

- Python 3.7+
- `requests`
- `msal`

## Usage

Update the TENANT_ID, CLIENT_ID, and CLIENT_SECRET in the script. Input any sharing URLs you would like to test. Example:
```bash
shared_urls = [
    "https://yourdomain.sharepoint.com/:t:/s/ExampleSite/ExampleFile.txt",
    "https://yourdomain.sharepoint.com/sites/yoursite/Lists/yourlistitem"
]

```

Then run:

```bash
python testShares.py
```

Install dependencies:
```bash
pip install requests msal
```


