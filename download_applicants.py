#!/usr/bin/env python3
"""
Script to download and backup the Applicants folder and related XML/CSV files
"""

import os
import shutil
import datetime
import zipfile
import argparse
from pathlib import Path

def create_backup(target_dir=None, include_credentials=False):
    """
    Create a backup of the Applicants folder and important data files.
    
    Args:
        target_dir (str): Target directory to save the backup to.
                         If None, saves to the current directory.
        include_credentials (bool): Whether to include credential files in the backup.
    
    Returns:
        str: Path to the created backup zip file
    """
    # Get current date for the backup filename
    current_date = datetime.datetime.now().strftime("%Y-%m-%d")
    
    # Determine target directory
    if target_dir is None:
        target_dir = os.getcwd()
    else:
        # Create target directory if it doesn't exist
        os.makedirs(target_dir, exist_ok=True)
    
    # Define backup filename
    backup_filename = os.path.join(target_dir, f"applicants_backup_{current_date}.zip")
    
    # Files to always include in backup
    files_to_backup = [
        "applicant_bank.xml",
        "applicant_bank_backup.xml",
        "applicants.xml",
        "applicants.csv",
        "last_run.json"
    ]
    
    # Credential files to optionally include
    credential_files = []
    if include_credentials:
        credential_files = [
            ".env",
            "service_account.json",
            "token_cache.json",
            "token_user_cache.json"
        ]
    
    # Create a zip file
    print(f"Creating backup at {backup_filename}")
    with zipfile.ZipFile(backup_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Add individual files
        for file in files_to_backup + credential_files:
            if os.path.exists(file):
                print(f"Adding file: {file}")
                zipf.write(file)
            else:
                print(f"Warning: File {file} not found, skipping")
        
        # Add the entire Applicants directory
        if os.path.exists("Applicants"):
            print("Adding Applicants folder...")
            
            # Walk through the directory and add all files
            for root, _, files in os.walk("Applicants"):
                for file in files:
                    file_path = os.path.join(root, file)
                    print(f"Adding: {file_path}")
                    zipf.write(file_path)
    
    print(f"\nBackup completed successfully: {backup_filename}")
    return backup_filename

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Create a backup of the Applicants folder and data files")
    parser.add_argument("--target", "-t", help="Target directory to save the backup to")
    parser.add_argument("--credentials", "-c", action="store_true", 
                       help="Include credential files in backup (not recommended for security reasons)")
    
    args = parser.parse_args()
    
    backup_file = create_backup(args.target, args.credentials)
    
    print(f"\nYou can download this file to your local machine.")
    print(f"If running in VS Code Server: Right-click on {backup_file} in the explorer and select 'Download'")