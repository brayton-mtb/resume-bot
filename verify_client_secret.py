import logging
import json
import os
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.auth.authentication_context import AuthenticationContext
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Get credentials from environment variables
SITE_URL = os.environ.get("SHAREPOINT_SITE_URL", "https://aheadcomputinginc.sharepoint.com/sites/ACHR")
CLIENT_ID = os.environ.get("MS_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("SHAREPOINT_CLIENT_SECRET", "")
TENANT_ID = os.environ.get("MS_TENANT_ID", "")
AUTHORITY_URL = f"https://login.microsoftonline.com/{TENANT_ID}"  # Construct the authority URL

def verify_client_secret(site_url, client_id, client_secret):
    credentials = ClientCredential(client_id, client_secret)
    ctx = ClientContext(site_url).with_credentials(credentials)
    try:
        # Attempt to fetch the SharePoint site details
        logging.debug("Attempting to fetch SharePoint site details...")
        ctx.web.get().execute_query()
        print("✅ Client secret is valid. Authentication successful!")
    except Exception as e:
        logging.error("❌ Authentication failed. Full error details:", exc_info=True)

        # Log the request and response details for debugging
        if hasattr(ctx.pending_request(), 'current_request'):
            request = ctx.pending_request().current_request
            logging.debug(f"Request URL: {request.url}")
            logging.debug(f"Request Headers: {json.dumps(dict(request.headers), indent=2)}")
            if request.data:
                logging.debug(f"Request Body: {request.data}")

        if hasattr(ctx.pending_request(), 'response'):
            response = ctx.pending_request().response
            logging.debug(f"Response Status Code: {response.status_code}")
            logging.debug(f"Response Headers: {json.dumps(dict(response.headers), indent=2)}")
            try:
                logging.debug(f"Response Body: {response.text}")
            except Exception as parse_error:
                logging.error(f"Failed to parse response body: {parse_error}")

        print(f"❌ Authentication failed: {e}")

# Function to fetch and print the token for debugging
def fetch_token(site_url, client_id, client_secret):
    try:
        auth_ctx = AuthenticationContext(AUTHORITY_URL)
        result = auth_ctx.acquire_token_for_app(client_id, client_secret)
        
        # Try to access the token based on how the library works
        if result:
            print("✅ Token acquired successfully!")
            # Try to inspect the result to see what's returned
            print("Token result type:", type(result))
            print("Token result:", result)
            
            # For office365 library, sometimes the token is the result itself
            if isinstance(result, str):
                print("Access Token:", result)
                return result
                
            # Try different common attribute/key names
            for attr in ['access_token', 'accessToken', 'token']:
                if hasattr(result, attr):
                    token = getattr(result, attr)
                    print(f"Access Token (from {attr}):", token)
                    return token
                elif isinstance(result, dict) and attr in result:
                    token = result[attr]
                    print(f"Access Token (from {attr}):", token)
                    return token
                    
            print("⚠️ Token acquired but unable to extract token string")
            return result  # Return whatever we got
        else:
            print("❌ Failed to acquire token")
            return None
    except Exception as e:
        print(f"❌ Error fetching token: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    print("🔄 Fetching token for debugging...")
    result = fetch_token(SITE_URL, CLIENT_ID, CLIENT_SECRET)
    print(result)
    verify_client_secret(SITE_URL, CLIENT_ID, CLIENT_SECRET)