#!/usr/bin/env python3
import imaplib
import email
import os
import json
import hashlib
from email.header import decode_header
from email.utils import parsedate_to_datetime
from datetime import datetime, timezone
from xml.etree.ElementTree import Element, SubElement, ElementTree, indent
import xml.etree.ElementTree as ET

IMAP_SERVER   = os.environ["IMAP_SERVER"]
IMAP_PORT     = int(os.environ.get("IMAP_PORT", 993))
EMAIL_USER    = os.environ["EMAIL_USER"]
EMAIL_PASS    = os.environ["EMAIL_PASS"]
SENDER_FILTER = os.environ["SENDER_FILTER"]
PAGES_URL     = os.environ["PAGES_URL"]
MAX_ITEMS     = int(os.environ.get("MAX_ITEMS", 100))
HISTORY_FILE  = "docs/history.json"
FEED_FILE     = "docs/feed.xml"

def decode_str(value):
    if value is None:
        return ""
    parts = decode_header(value)
    result = []
    for part, enc in parts:
        if isinstance(part, bytes):
            result.append(part.decode(enc or "utf-8", errors="replace"))
        else:
            result.append(part)
    return "".join(result)

def get_body(msg):
    if msg.is_multipart():
        html_body = None
        text_body = None
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get("Content-Disposition", ""))
            if "attachment" in cd:
                continue
            if ct == "text/html" and html_body is None:
                html_body = part.get_payload(decode=True).decode(
                    part.get_content_charset() or "utf-8", errors="replace")
            elif ct == "text/plain" and text_body is None:
                text_body = part.get_payload(decode=True).decode(
                    part.get_content_charset() or "utf-8", errors="replace")
        return html_body or text_body or ""
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            return payload.decode(msg.get_content_charset() or "utf-8", errors="replace")
        return ""

def fetch_emails():
    print(f"Connexion a {IMAP_SERVER}:{IMAP_PORT}...")
    conn = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    conn.login(EMAIL_USER, EMAIL_PASS)
    conn.select("INBOX", readonly=True)
    search_criteria = f'FROM "{SENDER_FILTER}"'
    _, data = conn.search(None, search_criteria)
    ids = data[0].split()
    print(f"{len(ids)} email(s) trouve(s) de {SENDER_FILTER}")
    items = []
    for uid in ids[-MAX_ITEMS:]:
        _, msg_data = conn.fetch(uid, "(RFC822)")
        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)
        subject = decode_str(msg.get("Subject", "(Sans objet)"))
        date_str = msg.get("Date", "")
        try:
            pub_date = parsedate_to_datetime(date_str)
        except Exception:
            pub_date = datetime.now(timezone.utc)
        body = get_body(msg)
        guid = hashlib.md5(f"{subject}{date_str}".encode()).hexdigest()
        items.append({"guid": guid, "title": subject, "pub_date": pub_date.isoformat(), "body": body})
    conn.logout()
    items.sort(key=lambda x: x["pub_date"], reverse=True)
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
    seen = {item["guid"] for item in existing}
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
    SubElement(channel, "description").text = f"Flux RSS auto-genere depuis les emails de {SENDER_FILTER}"
    SubElement(channel, "language").text = "fr"
    SubElement(channel, "lastBuildDate").text = datetime.now(timezone.utc).strftime("%a, %d %b %Y %H:%M:%S +0000")
    for item in items:
        entry = SubElement(channel, "item")
        SubElement(entry, "title").text = item["title"]
        SubElement(entry, "guid", isPermaLink="false").text = item["guid"]
        SubElement(entry, "link").text = f"{PAGES_URL}/feed.xml"
        try:
            dt = datetime.fromisoformat(item["pub_date"])
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
    new_items = fetch_emails()
    history   = load_history()
    all_items = merge_items(history, new_items)
    save_history(all_items)
    generate_rss(all_items)
    print("Termine.")
