#!/usr/bin/env python3
"""
TeacherOn Live Scraper & Verification Pipeline
Author: BMP Tutors Research Team
Purpose: Scrape and verify tutor profile URLs from the 'not scraped' cohort,
         extract inner HTML details, and update live status in SQLite database.

Usage:
    python3 scrape_unscraped_teacheron.py --limit 100 --delay 2.0
    python3 scrape_unscraped_teacheron.py --tutor-id T76707
    python3 scrape_unscraped_teacheron.py --proxy "http://user:pass@proxy:port"
"""

import sys
import os
import time
import argparse
import sqlite3
import random
import re
import urllib.request
import urllib.error
import ssl
from datetime import datetime

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:124.0) Gecko/20100101 Firefox/124.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_4_1) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4.1 Safari/605.1.15"
]

def parse_teacheron_html(html, url):
    """
    Parse TeacherOn tutor HTML profile page.
    Returns dictionary with parsed fields.
    """
    data = {
        "status": "active",
        "validation_reason": "Verified Live via Scraper",
        "name": None,
        "fee_text": None,
        "exp_text": None,
        "last_login": None,
        "registered": None
    }

    if "Attention Required! | Cloudflare" in html or "cf-browser-verification" in html:
        return {
            "status": "cloudflare_blocked",
            "validation_reason": "Cloudflare Challenge Triggered"
        }

    if "Page Not Found" in html or "404" in html or "This tutor is no longer available" in html:
        return {
            "status": "removed",
            "validation_reason": "No Profile on TeacherOn (404/Removed)"
        }

    # Extract name
    name_match = re.search(r'<h1[^>]*class="[^"]*tutor-name[^"]*"[^>]*>(.*?)</h1>', html, re.IGNORECASE | re.DOTALL)
    if not name_match:
        name_match = re.search(r'<h1[^>]*>(.*?)</h1>', html, re.DOTALL)
    if name_match:
        data["name"] = re.sub(r'<[^>]+>', '', name_match.group(1)).strip()

    # Extract Last Login
    login_match = re.search(r'Last login:\s*</span>\s*<span[^>]*>(.*?)</span>', html, re.IGNORECASE)
    if login_match:
        data["last_login"] = re.sub(r'<[^>]+>', '', login_match.group(1)).strip()

    # Extract Registered Date
    reg_match = re.search(r'Registered:\s*</span>\s*<span[^>]*>(.*?)</span>', html, re.IGNORECASE)
    if reg_match:
        data["registered"] = re.sub(r'<[^>]+>', '', reg_match.group(1)).strip()

    # Determine Active vs Dormant
    if data["registered"] and data["last_login"]:
        if data["registered"] == data["last_login"]:
            data["status"] = "dormant"
            data["validation_reason"] = "Inactive on TeacherOn (Dormant Single Registration)"
        else:
            data["status"] = "active"
            data["validation_reason"] = "Verified Active on TeacherOn"

    return data

def scrape_tutor_url(url, proxy=None, timeout=12, cookie=None):
    """
    Fetch a single TeacherOn tutor profile with browser spoofing.
    """
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    headers = {
        "User-Agent": random.choice(USER_AGENTS),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer": "https://www.teacheron.com/tutors-in-india",
        "DNT": "1",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
        "Upgrade-Insecure-Requests": "1"
    }

    if cookie:
        cookie_clean = cookie.strip()
        if "cf_clearance=" not in cookie_clean and "=" not in cookie_clean:
            cookie_clean = f"cf_clearance={cookie_clean}"
        headers["Cookie"] = cookie_clean

    req = urllib.request.Request(url, headers=headers)
    if proxy:
        req.set_proxy(proxy, "http" if proxy.startswith("http://") else "https")

    try:
        with urllib.request.urlopen(req, context=ctx, timeout=timeout) as resp:
            html = resp.read().decode("utf-8", errors="ignore")
            return parse_teacheron_html(html, url)
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return {"status": "removed", "validation_reason": "No Profile on TeacherOn (404/Removed)"}
        elif e.code == 403:
            return {"status": "cloudflare_blocked", "validation_reason": "Cloudflare 403 Forbidden"}
        else:
            return {"status": "http_error", "validation_reason": f"HTTP {e.code}"}
    except Exception as ex:
        return {"status": "error", "validation_reason": str(ex)[:40]}

def run_scraper(db_path, limit=50, delay=2.0, source_tab="not scraped", tutor_id=None, proxy=None, cookie=None):
    if not os.path.exists(db_path):
        print(f"Error: Database not found at {db_path}")
        sys.exit(1)

    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    if tutor_id:
        cur.execute("SELECT tutor_id, profile_url, teacher_name, registered_date, last_login_date, validation_reason FROM teachers WHERE tutor_id = ?", (tutor_id,))
    else:
        cur.execute("SELECT tutor_id, profile_url, teacher_name, registered_date, last_login_date, validation_reason FROM teachers WHERE source_tab = ? LIMIT ?", (source_tab, limit))

    rows = cur.fetchall()
    print(f"\n=======================================================")
    print(f" BMP Tutors - Live Profile Verification & Scraper")
    print(f" Target Cohort: {source_tab} | Queue Size: {len(rows)} tutors")
    print(f" Rate Limit Delay: {delay}s | Database: {os.path.basename(db_path)}")
    if cookie:
        print(f" Cloudflare Session Cookie: Configured (Length: {len(cookie)} chars)")
    print(f"=======================================================\n")

    stats = {"active": 0, "dormant": 0, "removed": 0, "cloudflare_blocked": 0, "errors": 0}

    for i, (tid, url, name, reg, last_log, current_val) in enumerate(rows, 1):
        if not url:
            continue

        print(f"[{i}/{len(rows)}] Scraping {tid} ({name or 'Anonymous'}) -> {url} ...", end=" ", flush=True)
        res = scrape_tutor_url(url, proxy=proxy, cookie=cookie)
        status = res.get("status")

        if status == "active":
            stats["active"] += 1
            print("✅ ACTIVE")
            cur.execute("UPDATE teachers SET is_active = 1, validation_reason = ? WHERE tutor_id = ?", (res["validation_reason"], tid))
        elif status == "dormant":
            stats["dormant"] += 1
            print("⏸️ DORMANT")
            cur.execute("UPDATE teachers SET is_active = 0, validation_reason = ? WHERE tutor_id = ?", (res["validation_reason"], tid))
        elif status == "removed":
            stats["removed"] += 1
            print("❌ 404 / REMOVED")
            cur.execute("UPDATE teachers SET is_active = 0, validation_reason = ? WHERE tutor_id = ?", (res["validation_reason"], tid))
        elif status == "cloudflare_blocked":
            stats["cloudflare_blocked"] += 1
            # Fall back to dataset timestamp verification
            if reg and last_log and reg != last_log:
                print("🛡️ CLOUDFLARE 403 (Verified Active via Saved Timestamps)")
            else:
                print("🛡️ CLOUDFLARE 403 (Bot Shield Active)")
        else:
            stats["errors"] += 1
            print(f"⚠️ {res.get('validation_reason')}")

        if i % 10 == 0:
            conn.commit()

        time.sleep(delay + random.uniform(0.1, 0.5))

    conn.commit()
    conn.close()

    print(f"\n=======================================================")
    print(f" Scraping & Verification Run Completed")
    print(f" Verified Active: {stats['active']} | Dormant: {stats['dormant']}")
    print(f" Profile Removed (404): {stats['removed']} | Cloudflare Shielded: {stats['cloudflare_blocked']}")
    print(f"=======================================================\n")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Scrape and verify TeacherOn tutor profiles")
    parser.add_argument("--db", default="cleaned_tutors_data.db", help="Path to SQLite database")
    parser.add_argument("--limit", type=int, default=10, help="Number of profiles to scrape")
    parser.add_argument("--delay", type=float, default=2.0, help="Polite delay between requests in seconds")
    parser.add_argument("--source-tab", default="not scraped", help="Cohort to scrape (not scraped / scraped / all)")
    parser.add_argument("--tutor-id", help="Single tutor ID to verify (e.g. T76707)")
    parser.add_argument("--proxy", help="HTTP/HTTPS proxy URL")

    args = parser.parse_args()
    run_scraper(
        db_path=args.db,
        limit=args.limit,
        delay=args.delay,
        source_tab=args.source_tab,
        tutor_id=args.tutor_id,
        proxy=args.proxy
    )
