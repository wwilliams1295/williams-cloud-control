print(">>> Script started")
from google.oauth2 import service_account
from googleapiclient.discovery import build
import os, json, base64, requests, msal
from dotenv import load_dotenv

load_dotenv(dotenv_path="/Users/williamwilliams/jarvis-demo/.env", override=True)

def test_google_drive():
    print("\n🔹 Testing Google Drive...")
    creds_json = os.getenv("GOOGLE_SA_JSON")

    if not creds_json:
        print("❌ No GOOGLE_SA_JSON found in environment.")
        return

    # Handle either JSON or base64 formats
    try:
        if creds_json.strip().startswith("{"):
            info = json.loads(creds_json)
        else:
            decoded = base64.b64decode(creds_json).decode()
            info = json.loads(decoded)
    except Exception as e:
        print("❌ Failed to parse GOOGLE_SA_JSON:", e)
        return

    try:
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/drive.readonly"]
        )
        drive = build("drive", "v3", credentials=creds)
        results = drive.files().list(pageSize=5, fields="files(id, name)").execute()
        files = results.get("files", [])

        if not files:
            print("ℹ️ No files visible to the service account (maybe folder not shared).")
        for f in files:
            print(f"✅ Google Drive file: {f['name']} ({f['id']})")
    except Exception as e:
        print("❌ Google Drive error:", e)

def test_onedrive():
    print("\n🔹 Testing OneDrive...")
    client_id  = os.environ["AZURE_CLIENT_ID"]
    tenant     = os.environ.get("AZURE_TENANT_ID", "common")
    cache_path = os.environ.get("ONEDRIVE_TOKEN_CACHE", ".onedrive_token.json")
    scopes     = ["Files.ReadWrite.All"]  # ✅ only this — others are reserved

    cache = msal.SerializableTokenCache()
    if os.path.exists(cache_path):
        try:
            cache.deserialize(open(cache_path, "r").read())
        except Exception as e:
            print(f"⚠️ Could not read cache {cache_path}: {e}")

    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant}",
        token_cache=cache
    )

    # Try silent first (requires cached account)
    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])

    # Fallback: device login (interactive once, local only)
    if not result or "access_token" not in result:
        print("ℹ️ No cached token found—starting device login (local, one-time).")
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"Failed to start device flow: {flow}")
        print(flow["message"])  # Follow the URL/code shown in terminal
        result = app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            raise RuntimeError(f"❌ OneDrive login failed: {result}")

        # Save refreshed cache
        try:
            open(cache_path, "w").write(cache.serialize())
            print(f"✅ Saved token cache → {cache_path}")
        except Exception as e:
            print(f"⚠️ Could not write cache {cache_path}: {e}")

    # Call Graph API
    r = requests.get(
        "https://graph.microsoft.com/v1.0/me/drive/root/children",
        headers={"Authorization": f"Bearer {result['access_token']}"}
    )

    print(f"ℹ️ Graph response code: {r.status_code}")
    if r.status_code == 200:
        items = r.json().get("value", [])
        if not items:
            print("ℹ️ OneDrive is reachable but the root folder is empty.")
        for i in items[:5]:
            print(f"✅ OneDrive file: {i['name']}")
    else:
        print("❌ OneDrive error:", r.status_code, r.text[:500])
if __name__ == "__main__":
    print(">>> Running integration tests...\n")
    test_google_drive()
    test_onedrive()
    print("\n>>> All tests completed.")
