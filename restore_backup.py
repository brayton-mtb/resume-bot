#!/usr/bin/env python3
"""
Script to restore a backup file from Google Drive, OneDrive, or local file
and extract it into the current workspace.

This script is designed to be used in your other VS Code workspace to import the data.
"""

import os
import sys
import argparse
import zipfile
import tempfile
import shutil
import datetime

def download_from_gdrive(file_id):
    """Download a file from Google Drive using the file ID"""
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload
        from google.oauth2 import service_account
        
        # Check if service_account.json exists
        if not os.path.exists("service_account.json"):
            print("❌ service_account.json not found. Please create this file with your Google API credentials.")
            return None
        
        print("🔑 Authenticating with Google Drive...")
        credentials = service_account.Credentials.from_service_account_file(
            'service_account.json',
            scopes=['https://www.googleapis.com/auth/drive.readonly']
        )
        
        service = build('drive', 'v3', credentials=credentials)
        
        # Create a temporary file to save the download
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.zip')
        temp_file.close()
        
        request = service.files().get_media(fileId=file_id)
        
        with open(temp_file.name, 'wb') as f:
            downloader = MediaIoBaseDownload(f, request)
            done = False
            print(f"📥 Downloading from Google Drive...")
            while not done:
                status, done = downloader.next_chunk()
                print(f"  Progress: {int(status.progress() * 100)}%")
        
        print(f"✅ Downloaded backup to {temp_file.name}")
        return temp_file.name
    except ImportError:
        print("❌ Google Drive download requires google-api-python-client and google-auth-httplib2 packages.")
        print("   Install them with: pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib")
        return None
    except Exception as e:
        print(f"❌ Google Drive download failed: {e}")
        return None

def extract_backup(backup_file, target_dir=None):
    """Extract the backup zip file to the current directory or specified target"""
    try:
        if target_dir is None:
            target_dir = os.getcwd()
        else:
            # Create target directory if it doesn't exist
            os.makedirs(target_dir, exist_ok=True)
        
        print(f"📂 Extracting backup to {target_dir}...")
        
        with zipfile.ZipFile(backup_file, 'r') as zip_ref:
            # Get list of all files in the zip
            file_list = zip_ref.namelist()
            total_files = len(file_list)
            
            # Extract each file with a progress indicator
            for i, file in enumerate(file_list):
                if i % 10 == 0 or i == total_files - 1:  # Show progress every 10 files
                    progress = (i + 1) / total_files * 100
                    print(f"  Progress: {progress:.1f}% ({i+1}/{total_files})", end='\r')
                
                # Extract the file, preserving directory structure
                zip_ref.extract(file, target_dir)
            
            print("\n✅ Extraction completed successfully!")
            
            # List the key files and directories that were restored
            print("\n📋 Restored files and directories:")
            root_items = set()
            for item in file_list:
                root_part = item.split('/')[0] if '/' in item else item
                root_items.add(root_part)
            
            for item in sorted(root_items):
                print(f"  - {item}")
        
        return True
    except Exception as e:
        print(f"❌ Extraction failed: {e}")
        return False

def backup_existing_data():
    """Create a backup of the current workspace data before overwriting"""
    try:
        # Check if important folders/files exist that should be backed up
        important_items = ['Applicants', 'applicants.xml', 'applicant_bank.xml', 'applicants.csv']
        items_to_backup = [item for item in important_items if os.path.exists(item)]
        
        if not items_to_backup:
            print("ℹ️ No existing data to back up.")
            return True
        
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        backup_dir = f"workspace_backup_{timestamp}"
        os.makedirs(backup_dir, exist_ok=True)
        
        print(f"📦 Backing up existing workspace data to {backup_dir}...")
        
        for item in items_to_backup:
            dest = os.path.join(backup_dir, item)
            print(f"  - {item}")
            if os.path.isdir(item):
                shutil.copytree(item, dest)
            else:
                shutil.copy2(item, dest)
        
        print("✅ Workspace backup completed successfully!")
        return True
    except Exception as e:
        print(f"❌ Workspace backup failed: {e}")
        return False

def main():
    parser = argparse.ArgumentParser(description="Restore a backup file into the current workspace")
    parser.add_argument("--source", choices=["gdrive", "file"], required=True,
                      help="Source of the backup: Google Drive or local file")
    parser.add_argument("--file-id", help="Google Drive file ID for the backup")
    parser.add_argument("--file", help="Local backup file path")
    parser.add_argument("--target", help="Target directory to extract to (defaults to current directory)")
    parser.add_argument("--skip-backup", action="store_true", help="Skip backing up existing workspace data")
    
    args = parser.parse_args()
    
    # Perform backup of existing workspace data
    if not args.skip_backup:
        if not backup_existing_data():
            print("⚠️ Warning: Failed to back up existing data. Continue anyway? (y/n)")
            response = input().strip().lower()
            if response != 'y':
                print("❌ Restoration aborted.")
                return 1
    
    # Get the backup file
    backup_file = None
    if args.source == "gdrive":
        if not args.file_id:
            print("❌ Google Drive source requires --file-id parameter")
            return 1
        backup_file = download_from_gdrive(args.file_id)
    elif args.source == "file":
        if not args.file:
            print("❌ File source requires --file parameter")
            return 1
        if not os.path.exists(args.file):
            print(f"❌ Backup file not found: {args.file}")
            return 1
        backup_file = args.file
    
    if not backup_file:
        print("❌ Failed to get backup file")
        return 1
    
    # Extract the backup
    if extract_backup(backup_file, args.target):
        if args.source == "gdrive":
            # Clean up the temporary file if we downloaded from Google Drive
            try:
                os.unlink(backup_file)
            except:
                pass
        print("✨ Restoration completed successfully!")
        return 0
    else:
        print("❌ Restoration failed")
        return 1

if __name__ == "__main__":
    sys.exit(main())