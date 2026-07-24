#!/usr/bin/env python3
import os, json, sys, time, requests, msal
from xml.etree.ElementTree import Element, SubElement, ElementTree, indent

CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
TENANT_ID     = os.environ.get("AZURE_TENANT_ID", "common")
REFRESH_TOKEN = os.environ["MS_REFRESH_TOKEN"]
SENDER_FILTER = os.environ["SENDER_FILTER"].lower()
PAGES_URL     = os.environ["PAGES_URL"]
MAX_ITEMS     = int(os.environ.get("MAX_ITEMS", 100))
HISTORY_FILE  = "docs/history.json"
FEED_FILE     = "docs/feed.xml"
SCOPES        = ["https://graph.microsoft.com/Mail.Read"]

def get_access_token():
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    )
    for attempt in range(1, 4):
        result = app.acquire_token_by_refresh_token(REFRESH_TOKEN, scopes=SCOPES)
        if "access_token" in result:
            print("[OK] Token acquired")
            return result["access_token"]
        print(f"[MSAL retry {attempt}/3]", json.dumps(result, indent=2), file=sys.stderr)
        if attempt < 3:
            time.sleep(5 * attempt)
    print("[MSAL] Token indisponible apres 3 tentatives - run ignore, feed conserve.", file=sys.stderr)
    return None

def fetch_emails(token):
    headers = {"Authorization": f"Bearer {token}"}
    # Fetch recent emails without OData filter (filter client-side)
    params = {
        "$orderby": "receivedDateTime desc",
        "$top": "100",
        "$select": "id,subject,receivedDateTime,body,from",
    }
    url = "https://graph.microsoft.com/v1.0/me/messages"
    last = None
    for attempt in range(1, 4):
        try:
            r = requests.get(url, headers=headers, params=params, timeout=30)
            if r.status_code == 200:
                all_emails = r.json().get("value", [])
                # Filter by sender client-side
                filtered = [
                    e for e in all_emails
                    if e.get("from", {}).get("emailAddress", {}).get("address", "").lower() == SENDER_FILTER
                ]
                print(f"[OK] {len(all_emails)} emails fetched, {len(filtered)} from {SENDER_FILTER}")
                return filtered
            last = f"HTTP {r.status_code}: {r.text}"
        except requests.RequestException as e:
            last = str(e)
        print(f"[GRAPH retry {attempt}/3] {last}", file=sys.stderr)
        if attempt < 3:
            time.sleep(5 * attempt)
    print(f"[GRAPH] Echec apres 3 tentatives: {last} - run ignore, feed conserve.", file=sys.stderr)
    return None

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
    if token is None:
        print("Pas de token ce run : feed propre.")
        sys.exit(0)
    emails = fetch_emails(token)
    if emails is None:
        print("Pas de fetch ce run : feed propre.")
        sys.exit(0)
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
