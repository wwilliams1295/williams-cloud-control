#!/usr/bin/env python3
"""
Setup Gmail Credentials for Render Deployment
This script helps set up Gmail API credentials for the deployed application.
"""

import os
import json
from pathlib import Path

def setup_gmail_credentials():
    """Set up Gmail credentials from environment variables."""
    
    # Check if we're on Render
    if os.getenv('RENDER'):
        print("Running on Render - setting up Gmail credentials from environment variables")
        
        # Get credentials from environment variables
        client_secret_json = os.getenv('GMAIL_CLIENT_SECRET_JSON')
        token_json = os.getenv('GMAIL_TOKEN_JSON')
        
        if not client_secret_json or not token_json:
            print("❌ Gmail credentials not found in environment variables")
            print("Please set GMAIL_CLIENT_SECRET_JSON and GMAIL_TOKEN_JSON in Render")
            return False
        
        try:
            # Parse and save client_secret.json
            client_secret_data = json.loads(client_secret_json)
            with open('client_secret.json', 'w') as f:
                json.dump(client_secret_data, f)
            
            # Parse and save token.json
            token_data = json.loads(token_json)
            with open('token.json', 'w') as f:
                json.dump(token_data, f)
            
            print("✅ Gmail credentials set up successfully")
            return True
            
        except json.JSONDecodeError as e:
            print(f"❌ Error parsing Gmail credentials: {e}")
            return False
        except Exception as e:
            print(f"❌ Error setting up Gmail credentials: {e}")
            return False
    else:
        print("Running locally - using existing Gmail credentials")
        return True

if __name__ == "__main__":
    setup_gmail_credentials()
