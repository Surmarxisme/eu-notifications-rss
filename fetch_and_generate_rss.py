#!/usr/bin/env python3
import os, json, sys, requests, msal
from xml.etree.ElementTree import Element, SubElement, ElementTree, indent

CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
TENANT_ID     = os.environ.get("AZURE_TENANT_ID", "common")
REFRESH_TOKEN = os.environ["MS_REFRESH_TOKEN"]
SENDER_FILTER = os.environ["SENDER_FILTER"]
PAGES_URL     = os.environ["PAGES_URL"]
MAX_ITEMS     = int(os.environ.get("MAX_ITEMS", 100))
HISTORY_FILE  = "docs/history.json"
FEED_FILE     = "docs/feed.xml"
SCOPES        = ["https://graph.microsoft.com/Mail.Read"]

def get_access_token():
    # Use the tenant from the secret (must match the tenant where the refresh token was issued)
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    )
    result = app.acquire_token_by_refresh_token(REFRESH_TOKEN, scopes=SCOPES)
    if "access_token" not in result:
        print("[MSAL ERROR]", json.dumps(result, indent=2), file=sys.stderr)
        sys.exit(1)
    print("[OK] Token acquired")
    # Print the token audience for debugging
    try:
        import base64
        parts = result["access_token"].split(".")
        payload = parts[1] + "==" * (4 - len(parts[1]) % 4)
        decoded = json.loads(base64.b64decode(payload))
        print(f"[DEBUG] Token aud={decoded.get('aud')} scp={decoded.get('scp')} upn={decoded.get('upn')} unique_name={decoded.get('unique_name')}", file=sys.stderr)
    except Exception as e:
        print(f"[DEBUG] Could not decode token: {e}", file=sys.stderr)
    return result["access_token"]

def fetch_emails(token):
    headers = {"Authorization": f"Bearer {token}"}
    params = {
        "$filter": f"from/emailAddress/address eq '{SENDER_FILTER}'",
        "$orderby": "receivedDateTime desc",
        "$top": "100",
        "$select": "id,subject,receivedDateTime,body",
    }
    url = "https://graph.microsoft.com/v1.0/me/messages"
    r = requests.get(url, headers=headers, params=params)
    if r.status_code != 200:
        print(f"[GRAPH ERROR] {r.status_code}: {r.text}", file=sys.stderr)
        r.raise_for_status()
    return r.json().get("value", [])

def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE) as f:
            return json.load(f)
    return []

def save_history(history):
    os.makedirs("docs", exist_ok=True)
    with open(HISTORY_FILE, "w") as f:
        json.dump(history, f, indent=2)

def build_rss(items):
    rss = Element("rss", version="2.0")
    channel = SubElement(rss, "channel")
    SubElement(channel, "title").text = "EU Notifications"
    SubElement(channel, "link").text = PAGES_URL
    SubElement(channel, "description").text = "Emails from EU Corporate Notification System"
    for item in items[:MAX_ITEMS]:
        entry = SubElement(channel, "item")
        SubElement(entry, "title").text = item.get("subject", "")
        SubElement(entry, "pubDate").text = item.get("receivedDateTime", "")
        SubElement(entry, "guid").text = item.get("id", "")
        SubElement(entry, "description").text = item.get("body", {}).get("content", "")
    indent(rss)
    return ElementTree(rss)

if __name__ == "__main__":
    token = get_access_token()
    emails = fetch_emails(token)
    print(f"[OK] Fetched {len(emails)} emails")
    history = load_history()
    seen_ids = {e["id"] for e in history}
    new_items = [e for e in emails if e["id"] not in seen_ids]
    print(f"[OK] {len(new_items)} new items")
    all_items = new_items + history
    save_history(all_items[:MAX_ITEMS])
    tree = build_rss(all_items)
    os.makedirs("docs", exist_ok=True)
    tree.write(FEED_FILE, encoding="unicode", xml_declaration=True)
    print(f"[OK] Feed written to {FEED_FILE}")
