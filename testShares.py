import os
import json
import time
import requests
import msal
import base64

# === CONFIGURATION ===
TENANT_ID = 'YOUR_TENANT_ID'
CLIENT_ID = 'YOUR_CLIENT_ID'
CLIENT_SECRET = 'YOUR_CLIENT_SECRET'
AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'
SCOPE = ['https://graph.microsoft.com/.default']

# === EXAMPLE USAGE ===
shared_urls = [
    "https://yourdomain.sharepoint.com/:t:/s/ExampleSite/ExampleFile.txt",
    "https://yourdomain.sharepoint.com/:li:/s/ExampleSite/ExampleListItem",
    "https://yourdomain.sharepoint.com/sites/yoursite/Shared%20Documents/yourfile.docx",
    "https://yourdomain.sharepoint.com/sites/yoursite/Lists/yourlistitem"
]

# === ACQUIRE ACCESS TOKEN ===
def get_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=SCOPE)
    if 'access_token' in result:
        return result['access_token']
    else:
        raise Exception("Could not obtain access token")

# === HANDLE API REQUESTS WITH RETRY ===
def make_api_call(url, headers, method='GET', data=None):
    while True:
        try:
            if method == 'GET':
                response = requests.get(url, headers=headers)
            elif method == 'POST':
                response = requests.post(url, headers=headers, data=data)
            elif method == 'PUT':
                response = requests.put(url, headers=headers, data=data)
            else:
                raise ValueError("Unsupported HTTP method")

            if response.status_code == 429:
                retry_after = int(response.headers.get('Retry-After', 1))
                print(f"Throttled. Retrying after {retry_after} seconds.")
                time.sleep(retry_after)
            else:
                response.raise_for_status()
                return response
        except requests.exceptions.RequestException as e:
            print(f"Request failed: {e}. Retrying in 10 seconds...")
            time.sleep(10)

# === ENCODE SHARED URL TO SHAREID ===
def encode_share_url(url):
    encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
    return f"u!{encoded}"

# === DECODE SHAREID AND CLASSIFY RESOURCE TYPE ===
def decode_share_id(share_id):
    if not share_id.startswith("u!"):
        return None, "Invalid shareId format"
    encoded_part = share_id[2:]
    padding = '=' * (-len(encoded_part) % 4)
    decoded_bytes = base64.urlsafe_b64decode(encoded_part + padding)
    decoded_url = decoded_bytes.decode("utf-8")

    # Classify based on known URL patterns
    if ":t:" in decoded_url:
        resource_type = "Document (Text)"
    elif ":w:" in decoded_url:
        resource_type = "Word Document"
    elif ":x:" in decoded_url:
        resource_type = "Excel Spreadsheet"
    elif ":p:" in decoded_url:
        resource_type = "PowerPoint Presentation"
    elif ":i:" in decoded_url:
        resource_type = "Image"
    elif ":v:" in decoded_url:
        resource_type = "Video"
    elif ":li:" in decoded_url:
        resource_type = "List Item"
    elif "/Shared Documents/" in decoded_url or decoded_url.endswith(('.docx', '.pdf', '.xlsx', '.pptx')):
        resource_type = "Document"
    elif "/Lists/" in decoded_url:
        resource_type = "List Item"
    else:
        resource_type = "Unknown or unsupported type"

    return decoded_url, resource_type

# === INSPECT SHARED URL AND CALL APPROPRIATE GRAPH ENDPOINT ===
def inspect_share_metadata(share_url, access_token):
    share_id = encode_share_url(share_url)
    decoded_url, resource_type = decode_share_id(share_id)

    print(f"\n--- Testing Shared URL ---")
    print(f"Decoded URL: {decoded_url}")
    print(f"Classified Resource Type: {resource_type}")

    base_endpoint = f"https://graph.microsoft.com/v1.0/shares/{share_id}"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
        "Prefer": "redeemSharingLinkIfNecessary"
    }

    # Request metadata
    metadata_response = make_api_call(base_endpoint, headers)
    print(f"Metadata Request URL: {base_endpoint}")
    print(f"Status Code: {metadata_response.status_code}")
    try:
        metadata_json = metadata_response.json()
        print("Metadata Response JSON:")
        print(json.dumps(metadata_json, indent=2))
    except Exception as e:
        print(f"Failed to parse metadata response: {e}")

    # Determine endpoint based on classification
    if "Document" in resource_type:
        item_endpoint = f"{base_endpoint}/driveItem"
    elif "List Item" in resource_type:
        item_endpoint = f"{base_endpoint}/listItem"
    else:
        print("Skipping item retrieval due to unknown resource type.")
        return

    # Request item details
    item_response = make_api_call(item_endpoint, headers)
    print(f"Item Request URL: {item_endpoint}")
    print(f"Status Code: {item_response.status_code}")
    try:
        item_json = item_response.json()
        print("Item Response JSON:")
        print(json.dumps(item_json, indent=2))
    except Exception as e:
        print(f"Failed to parse item response: {e}")

access_token = get_token()

for url in shared_urls:
    inspect_share_metadata(url, access_token)
