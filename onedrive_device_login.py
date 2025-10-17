import os, msal
from dotenv import load_dotenv

load_dotenv(dotenv_path="/Users/williamwilliams/jarvis-demo/.env", override=True)

CLIENT_ID  = os.environ["AZURE_CLIENT_ID"]
TENANT     = os.environ.get("AZURE_TENANT_ID", "common")
CACHE_PATH = os.environ.get("ONEDRIVE_TOKEN_CACHE", ".onedrive_token.json")
SCOPES = ["Files.ReadWrite.All"]


cache = msal.SerializableTokenCache()
app = msal.PublicClientApplication(
    CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT}",
    token_cache=cache
)

flow = app.initiate_device_flow(scopes=SCOPES)
if "user_code" not in flow:
    raise SystemExit(f"Failed to start device flow: {flow}")
print("\n== Device Login ==\n", flow["message"], "\n")

result = app.acquire_token_by_device_flow(flow)
if "access_token" not in result:
    raise SystemExit(f"Login failed: {result}")

open(CACHE_PATH, "w").write(cache.serialize())
print("✅ Saved token cache →", CACHE_PATH)
