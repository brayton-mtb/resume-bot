# resume-bot

A tool to filter and organize job applicant resumes.

# Setup Instructions

## Sensitive Credential Files
This project requires several credential files that are not included in the repository for security reasons. Before running the application, you need to create these files using the provided example templates:

1. **Service Account Credentials**:
   - Copy `service_account.json.example` to `service_account.json`
   - Replace the placeholder values with your actual Google Cloud service account credentials

2. **Microsoft Authentication Tokens**:
   - Copy `token_cache.json.example` to `token_cache.json`
   - Copy `token_user_cache.json.example` to `token_user_cache.json`
   - These files will be automatically populated when you run the application and authenticate

## Running the Application
1. Make sure you have all required dependencies installed: 
   ```
   pip install -r requirements.txt
   ```

2. Run the main application:
   ```
   python gen_filter_bot.py
   ```

## Development Notes
- The original `resume_filter_bot.py` is preserved for reference but should not be edited
- Use `gen_filter_bot.py` for all new development

# Preference saved: Never edit resume_filter_bot.py