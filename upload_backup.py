#!/usr/bin/env python3
"""
Script to upload the latest backup to an easily accessible location
This script supports multiple upload methods:
1. FTP - Upload to an FTP server
2. Google Drive - Upload to a Google Drive folder (using service_account.json)
3. Microsoft OneDrive - Upload to a OneDrive folder (using existing Microsoft auth)

Usage:
  python upload_backup.py --method [ftp|gdrive|onedrive] [--options]
"""

import os
import sys
import argparse
import datetime
import glob
from pathlib import Path

def find_latest_backup():
    """Find the latest backup file in the current directory"""
    backup_files = glob.glob("applicants_backup_*.zip")
    if not backup_files:
        print("❌ No backup files found. Please run download_applicants.py first.")
        sys.exit(1)
    
    # Sort by modification time (newest first)
    latest = max(backup_files, key=lambda f: os.path.getmtime(f))
    print(f"📦 Found latest backup: {latest}")
    return latest

def upload_ftp(backup_file, host, username, password, remote_dir=None):
    """Upload the backup file to an FTP server"""
    try:
        import ftplib
        
        print(f"📤 Connecting to FTP server {host}...")
        ftp = ftplib.FTP(host)
        ftp.login(username, password)
        
        if remote_dir:
            try:
                ftp.cwd(remote_dir)
            except ftplib.error_perm:
                # Try to create the directory if it doesn't exist
                print(f"📁 Creating directory {remote_dir}")
                ftp.mkd(remote_dir)
                ftp.cwd(remote_dir)
        
        with open(backup_file, 'rb') as file:
            print(f"📤 Uploading {backup_file}...")
            ftp.storbinary(f'STOR {os.path.basename(backup_file)}', file, callback=lambda _: print(".", end="", flush=True))
        
        print(f"\n✅ Successfully uploaded {backup_file} to FTP server {host}")
        ftp.quit()
        return True
    except ImportError:
        print("❌ FTP upload requires the ftplib module. It should be included in Python standard library.")
        return False
    except Exception as e:
        print(f"❌ FTP upload failed: {e}")
        return False

def upload_gdrive(backup_file, folder_id=None):
    """Upload the backup file to Google Drive"""
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        from google.oauth2 import service_account
        
        # Check if service_account.json exists
        if not os.path.exists("service_account.json"):
            print("❌ service_account.json not found. Please create this file with your Google API credentials.")
            return False
        
        print("🔑 Authenticating with Google Drive...")
        credentials = service_account.Credentials.from_service_account_file(
            'service_account.json',
            scopes=['https://www.googleapis.com/auth/drive']
        )
        
        service = build('drive', 'v3', credentials=credentials)
        
        file_metadata = {
            'name': os.path.basename(backup_file),
        }
        
        if folder_id:
            file_metadata['parents'] = [folder_id]
        
        media = MediaFileUpload(backup_file, resumable=True)
        
        print(f"📤 Uploading {backup_file} to Google Drive...")
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()
        
        print(f"✅ Successfully uploaded to Google Drive with file ID: {file.get('id')}")
        
        # Make the file publicly accessible with a shareable link
        try:
            permission = {
                'type': 'anyone',
                'role': 'reader'
            }
            service.permissions().create(
                fileId=file.get('id'),
                body=permission
            ).execute()
            
            # Get the web link to the file
            file_data = service.files().get(
                fileId=file.get('id'),
                fields='webViewLink'
            ).execute()
            
            print(f"🔗 File can be downloaded from: {file_data.get('webViewLink')}")
        except Exception as e:
            print(f"⚠️ Could not make file public: {e}")
        
        return True
    except ImportError:
        print("❌ Google Drive upload requires google-api-python-client and google-auth-httplib2 packages.")
        print("   Install them with: pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib")
        return False
    except Exception as e:
        print(f"❌ Google Drive upload failed: {e}")
        return False

def upload_onedrive(backup_file, folder_path=None):
    """Upload the backup file to Microsoft OneDrive"""
    try:
        import msal
        import requests
        import json
        import time
        
        # Check if token_cache.json exists
        if not os.path.exists("token_cache.json"):
            print("❌ token_cache.json not found. Please log in to Microsoft first.")
            return False
            
        # Load environment variables or config
        client_id = os.environ.get("MS_CLIENT_ID")
        tenant_id = os.environ.get("MS_TENANT_ID")
        
        if not client_id or not tenant_id:
            # Try to load from .env file
            try:
                from dotenv import load_dotenv
                load_dotenv()
                client_id = os.environ.get("MS_CLIENT_ID")
                tenant_id = os.environ.get("MS_TENANT_ID")
            except ImportError:
                pass
        
        if not client_id or not tenant_id:
            print("❌ Microsoft client ID or tenant ID not found in environment variables.")
            print("   Please set MS_CLIENT_ID and MS_TENANT_ID environment variables.")
            return False
        
        # Initialize MSAL authentication
        cache = msal.SerializableTokenCache()
        
        # Load token cache from file
        with open("token_cache.json", "r") as token_file:
            cache.deserialize(token_file.read())
        
        # Set up MSAL application
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app = msal.PublicClientApplication(
            client_id=client_id,
            authority=authority,
            token_cache=cache
        )
        
        # Get account and acquire token silently
        accounts = app.get_accounts()
        if accounts:
            scopes = ["Files.ReadWrite"]
            result = app.acquire_token_silent(scopes, account=accounts[0])
            
            if "access_token" in result:
                print("🔑 Successfully acquired Microsoft token from cache")
                
                # Upload the file
                token = result["access_token"]
                headers = {
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/octet-stream"
                }
                
                # Determine upload path
                if folder_path:
                    upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{folder_path}/{os.path.basename(backup_file)}:/content"
                else:
                    upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{os.path.basename(backup_file)}:/content"
                
                print(f"📤 Uploading {backup_file} to OneDrive...")
                
                with open(backup_file, "rb") as file:
                    response = requests.put(upload_url, headers=headers, data=file)
                
                if response.status_code in [200, 201]:
                    print("✅ Successfully uploaded to OneDrive")
                    
                    # Get sharable link
                    file_data = response.json()
                    file_id = file_data.get("id")
                    
                    share_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/createLink"
                    share_payload = {
                        "type": "view",
                        "scope": "anonymous"
                    }
                    
                    share_headers = {
                        "Authorization": f"Bearer {token}",
                        "Content-Type": "application/json"
                    }
                    
                    share_response = requests.post(
                        share_url, 
                        headers=share_headers, 
                        data=json.dumps(share_payload)
                    )
                    
                    if share_response.status_code == 200:
                        share_data = share_response.json()
                        share_link = share_data.get("link", {}).get("webUrl")
                        print(f"🔗 File can be downloaded from: {share_link}")
                    else:
                        print(f"⚠️ Could not create sharing link: {share_response.text}")
                    
                    return True
                else:
                    print(f"❌ Failed to upload to OneDrive: {response.status_code} {response.text}")
                    return False
            else:
                print("❌ Token acquisition failed. You may need to log in again.")
                return False
        else:
            print("❌ No Microsoft accounts found in cache.")
            return False
    except ImportError:
        print("❌ OneDrive upload requires msal and requests packages.")
        print("   Install them with: pip install msal requests")
        return False
    except Exception as e:
        print(f"❌ OneDrive upload failed: {e}")
        return False

def main():
    parser = argparse.ArgumentParser(description="Upload backup file to remote location")
    parser.add_argument("--method", choices=["ftp", "gdrive", "onedrive"], required=True,
                       help="Upload method: FTP, Google Drive or OneDrive")
    
    # FTP options
    parser.add_argument("--host", help="FTP server hostname")
    parser.add_argument("--user", help="FTP username")
    parser.add_argument("--password", help="FTP password")
    parser.add_argument("--dir", help="Remote directory on FTP server")
    
    # Google Drive options
    parser.add_argument("--folder-id", help="Google Drive folder ID to upload to")
    
    # OneDrive options
    parser.add_argument("--onedrive-folder", help="OneDrive folder path")
    
    args = parser.parse_args()
    
    # Find the latest backup
    backup_file = find_latest_backup()
    
    # Upload based on selected method
    if args.method == "ftp":
        if not args.host or not args.user or not args.password:
            print("❌ FTP upload requires --host, --user, and --password arguments")
            return 1
        
        if upload_ftp(backup_file, args.host, args.user, args.password, args.dir):
            return 0
        else:
            return 1
    
    elif args.method == "gdrive":
        if upload_gdrive(backup_file, args.folder_id):
            return 0
        else:
            return 1
    
    elif args.method == "onedrive":
        if upload_onedrive(backup_file, args.onedrive_folder):
            return 0
        else:
            return 1

if __name__ == "__main__":
    sys.exit(main())