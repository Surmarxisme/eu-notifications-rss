#!/usr/bin/env python3
import os, json, hashlib, requests, msal
from datetime import datetime, timezone
from xml.etree.ElementTree import Element, SubElement, ElementTree, indent
import xml.etree.ElementTree as ET

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
    app = msal.PublicClientApplication(CLIENT_ID,
          authority=f"https://login.microsoftonline.com/{TENANT_ID}")
    result = app.acquire_token_by_refresh_token(REFRESH_TOKEN, scopes=SCOPES)
    if "access_token" not in result:
        raise Exception(f"Token error: {result.get('error_description')}")
    return result["access_token"]

def fetch_emails(token):
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        "https://graph.microsoft.com/v1.0/me/messages"
        f"?$filter=from/emailAddress/address eq '{SENDER_FILTER}'"
        f"&$orderby=receivedDateTime desc"
        f"&$top={MAX_ITEMS}"
        f"&$select=id,subject,receivedDateTime,body"
    )
    items = []
    while url:
        r = requests.get(url, headers=headers)
        r.raise_for_status()
        data = r.json()
        for msg in data.get("value", []):
            guid = hashlib.md5(msg["id"].encode()).hexdigest()
            items.append({
                "guid": guid,
                "title": msg.get("subject", "(Sans objet)"),
                "pub_date": msg["receivedDateTime"],
                "body": msg["body"]["content"],
            })
        url = data.get("@odata.nextLink")
    print(f"{len(items)} email(s) recupere(s) de {SENDER_FILTER}")
    return items

def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_history(items):
    os.makedirs("docs", exist_ok=True)
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=2)

def merge_items(existing, new):
    seen = {i["guid"] for i in existing}
    merged = list(existing)
    added = 0
    for item in new:
        if item["guid"] not in seen:
            merged.append(item)
            seen.add(item["guid"])
            added += 1
    print(f"{added} nouveau(x) item(s) ajoute(s).")
    merged.sort(key=lambda x: x["pub_date"], reverse=True)
    return merged[:MAX_ITEMS]

def generate_rss(items):
    rss = Element("rss", version="2.0")
    rss.set("xmlns:content", "http://purl.org/rss/1.0/modules/content/")
    channel = SubElement(rss, "channel")
    SubElement(channel, "title").text = f"Notifications Commission UE - {SENDER_FILTER}"
    SubElement(channel, "link").text = PAGES_URL
    SubElement(channel, "description").text = f"Flux RSS auto-genere depuis {SENDER_FILTER}"
    SubElement(channel, "language").text = "fr"
    SubElement(channel, "lastBuildDate").text = datetime.now(timezone.utc).strftime("%a, %d %b %Y %H:%M:%S +0000")
    for item in items:
        entry = SubElement(channel, "item")
        SubElement(entry, "title").text = item["title"]
        SubElement(entry, "guid", isPermaLink="false").text = item["guid"]
        SubElement(entry, "link").text = f"{PAGES_URL}/feed.xml"
        try:
            dt = datetime.fromisoformat(item["pub_date"].replace("Z", "+00:00"))
            SubElement(entry, "pubDate").text = dt.strftime("%a, %d %b %Y %H:%M:%S +0000")
        except Exception:
            pass
        content_el = SubElement(entry, "{http://purl.org/rss/1.0/modules/content/}encoded")
        content_el.text = item["body"]
    indent(rss, space="  ")
    os.makedirs("docs", exist_ok=True)
    tree = ElementTree(rss)
    ET.register_namespace("content", "http://purl.org/rss/1.0/modules/content/")
    with open(FEED_FILE, "wb") as f:
        tree.write(f, encoding="utf-8", xml_declaration=True)
    print(f"Flux RSS ecrit : {FEED_FILE} ({len(items)} items)")

if __name__ == "__main__":
    token     = get_access_token()
    new_items = fetch_emails(token)
    history   = load_history()
    all_items = merge_items(history, new_items)
    save_history(all_items)
    generate_rss(all_items)
    print("Termine.")
