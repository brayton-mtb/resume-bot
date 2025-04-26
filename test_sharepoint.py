import os
import msal
import requests
import json
import logging
import webbrowser
import time
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Configure logging
logging.basicConfig(level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# SharePoint configuration
TENANT_ID = os.environ.get("MS_TENANT_ID", "")
CLIENT_ID = os.environ.get("MS_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("SHAREPOINT_CLIENT_SECRET", "")  # Get from environment variables
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SHAREPOINT_HOST = "aheadcomputinginc.sharepoint.com"
SITE_NAME = "ACHR"
RESOURCE = f"https://{SHAREPOINT_HOST}/"  # Must include trailing slash

# SharePoint resource ID is different from Microsoft Graph
SHAREPOINT_RESOURCE_ID = "00000003-0000-0ff1-ce00-000000000000"  # Correct SharePoint Resource ID

# These are the scopes needed for SharePoint access
SCOPES = [
    f"{SHAREPOINT_RESOURCE_ID}/.default"  # Request all delegated permissions for SharePoint
]

def get_user_interactive_token():
    """
    Obtains an access token using interactive authentication via device code flow.
    """
    # Load cache from file if it exists
    cache_file = "token_user_cache.json"
    cache = msal.SerializableTokenCache()
    
    if os.path.exists(cache_file):
        with open(cache_file, "r") as f:
            cache.deserialize(f.read())
            
    # Create public client application with token cache
    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        token_cache=cache
    )
    
    # Check if there's a token in the cache
    accounts = app.get_accounts()
    if accounts:
        logging.info(f"Found {len(accounts)} account(s) in cache. Attempting silent auth.")
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            logging.info("Successfully obtained token from cache.")
            # Save cache
            with open(cache_file, "w") as f:
                f.write(cache.serialize())
            return result["access_token"]
    
    # Fall back to device code flow if silent auth fails
    logging.info("Starting device code flow for authentication...")
    flow = app.initiate_device_flow(scopes=SCOPES)
    
    if "user_code" not in flow:
        logging.error(f"Failed to start device flow: {flow.get('error_description', flow.get('error'))}")
        return None
    
    # Display instructions for the user
    print("\n" + "="*50)
    print("🔐 AUTHENTICATION REQUIRED 🔐")
    print("="*50)
    print("\nTo connect to SharePoint, you need to sign in with your Microsoft account.")
    print(f"\n1. Go to: {flow['verification_uri']}")
    print(f"2. Enter this code: {flow['user_code']}")
    print(f"\n3. Sign in with your Microsoft account that has access to SharePoint.")
    print("\nThis authentication uses delegated permissions, which means:")
    print("- The app can only access what YOU have access to")
    print("- No admin consent should be required")
    print("="*50 + "\n")
    
    # Open the browser automatically
    try:
        webbrowser.open(flow["verification_uri"])
    except Exception as e:
        logging.warning(f"Could not open browser automatically: {e}")
    
    # Wait for the user to authenticate
    result = app.acquire_token_by_device_flow(flow)
    
    if "access_token" in result:
        # Save token cache
        with open(cache_file, "w") as f:
            f.write(cache.serialize())
            
        token = result["access_token"]
        logging.info("✅ Successfully obtained access token!")
        
        # Log token details (safely)
        token_parts = token.split('.')
        if len(token_parts) >= 2:
            try:
                import base64
                
                # Handle base64 padding
                def decode_base64(data):
                    padding = len(data) % 4
                    if padding:
                        data += '=' * (4 - padding)
                    return base64.b64decode(data)
                
                # Parse payload (claims)
                payload_json = decode_base64(token_parts[1])
                payload = json.loads(payload_json)
                
                # Log important claims for debugging
                safe_payload = {
                    "aud": payload.get("aud"),
                    "iss": payload.get("iss"),
                    "exp": payload.get("exp"),
                    "scp": payload.get("scp", ""),
                    "username": payload.get("upn", payload.get("preferred_username", "--")),
                    "expires": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(payload.get("exp", 0)))
                }
                logging.info(f"Token details: {json.dumps(safe_payload, indent=2)}")
            except Exception as e:
                logging.error(f"Failed to decode token: {e}")
        
        return token
    else:
        error = result.get("error", "unknown")
        error_desc = result.get("error_description", "No description")
        logging.error(f"❌ Failed to acquire token. Error: {error}. Description: {error_desc}")
        return None

def test_sharepoint_connection():
    """
    Tests the connection to SharePoint Online using interactive user authentication.
    """
    token = get_user_interactive_token()
    if not token:
        return
    
    print("\n✅ Successfully authenticated! Now testing SharePoint connection...\n")
    
    # Test endpoints
    endpoints = [
        f"https://{SHAREPOINT_HOST}/sites/{SITE_NAME}/_api/web",
        f"https://{SHAREPOINT_HOST}/sites/{SITE_NAME}/_api/contextinfo",  # This needs POST
    ]
    
    success = False
    
    for endpoint in endpoints:
        method = "POST" if "contextinfo" in endpoint else "GET"
        logging.info(f"Testing endpoint with {method}: {endpoint}")
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose"
        }
        
        try:
            if method == "GET":
                response = requests.get(endpoint, headers=headers)
            else:
                response = requests.post(endpoint, headers=headers)
            
            logging.info(f"Response status: {response.status_code}")
            
            if response.status_code == 200 or response.status_code == 201:
                success = True
                logging.info("✅ Request successful!")
                print(f"✅ Successfully connected to SharePoint endpoint: {endpoint}")
                
                # Log a sample of the response
                response_json = response.json()
                logging.debug(f"Response sample: {str(response_json)[:200]}...")
                
                # If this is the context info endpoint, extract form digest value
                if "contextinfo" in endpoint and response.status_code == 200:
                    try:
                        form_digest = response_json["d"]["GetContextWebInformation"]["FormDigestValue"]
                        logging.info(f"Form digest value: {form_digest}")
                        print(f"✅ Obtained form digest value, which confirms valid SharePoint access!")
                    except (KeyError, TypeError) as e:
                        logging.warning(f"Could not extract form digest value: {e}")
            else:
                logging.error(f"❌ Request failed with status {response.status_code}")
                print(f"❌ Failed to connect to endpoint {endpoint}: {response.status_code}")
                logging.error(f"Response: {response.text[:500]}")
                print(f"Error: {response.text}")
        except Exception as e:
            logging.error(f"❌ Request error: {e}")
            print(f"❌ Error connecting to SharePoint: {e}")
    
    if success:
        print("\n✅ AUTHENTICATION AND CONNECTION TEST SUCCESSFUL!")
        print("You can now modify your upload_folder_to_sharepoint function to use this authentication approach.")
        print("To update gen_filter_bot.py, use the user delegated authentication flow that worked in this test.")
    else:
        print("\n❌ CONNECTION TEST FAILED")
        print("SharePoint authentication was successful, but the connection test failed.")
        print("This might be due to permission issues or incorrect site URL.")

def main():
    print("\n🔄 Starting SharePoint connection test with user interactive authentication...")
    test_sharepoint_connection()
    print("\nTest completed.")

if __name__ == "__main__":
    main()