#!/usr/bin/env python3
"""
Verify scraped teacheron tutor profiles from scraped_export.csv.
Checks each profile_url to see if the tutor page is still active or removed.
Separate from verify.py (which handles the unscraped cohort unverified_31017_tutors.csv).

Usage:
    python3 verify_scraped_tutors.py

Input:
    scraped_export.csv  (columns: profile_url, Subjects, Experience, Education, Teacher_Name, ...)

Output:
    scraped_verification_results.csv  (tutor_id/profile_url/is_active/status_reason + original fields)
"""

import csv
import os
import time
import random
from playwright.sync_api import sync_playwright

INPUT_CSV = "retry_scraped.csv"
OUTPUT_CSV = "retry_scraped_results.csv"

USER_AGENT = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

# Load already-processed profile URLs so we can resume after interruptions
done_urls = set()
if os.path.exists(OUTPUT_CSV) and os.path.getsize(OUTPUT_CSV) > 0:
    with open(OUTPUT_CSV, "r", encoding="utf-8") as f:
        for line in f:
            parts = line.strip().split(",")
            if parts:
                done_urls.add(parts[0])

with sync_playwright() as p:
    # headless=False => visible browser so you can solve Cloudflare manually
    browser = p.chromium.launch(
        headless=False,
        args=[
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox",
            "--disable-dev-shm-usage",
        ],
    )
    context = browser.new_context(
        user_agent=USER_AGENT,
        viewport={"width": 1366, "height": 768},
    )
    page = context.new_page()

    page.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
    """)

    write_header = not os.path.exists(OUTPUT_CSV) or os.path.getsize(OUTPUT_CSV) == 0

    with (
        open(INPUT_CSV, "r", encoding="utf-8") as f_in,
        open(OUTPUT_CSV, "a", newline="", encoding="utf-8") as f_out,
    ):
        reader = csv.DictReader(f_in)

        # Build output header: original fields + verification fields
        original_fields = reader.fieldnames or []
        out_fields = original_fields + ["is_active", "status_reason"]
        writer = csv.DictWriter(f_out, fieldnames=out_fields, extrasaction="ignore")
        if write_header:
            writer.writeheader()

        for row in reader:
            url = row.get("profile_url", "").strip()
            if not url or url in done_urls:
                continue  # skip empty or already-processed

            try:
                res = page.goto(url, wait_until="networkidle", timeout=30000)

                try:
                    page.wait_for_selector("h1", timeout=5000)
                except Exception:
                    pass

                content = page.content()

                if (
                    "Sorry, you have been blocked" in content
                    or "Attention Required!" in content
                    or "Just a moment" in content
                ):
                    print(
                        f"\n[{url}] Cloudflare blocked! Solve the challenge in the browser window, then press ENTER here to continue..."
                    )
                    input()
                    content = page.content()
                    if (
                        "Sorry, you have been blocked" in content
                        or "Attention Required!" in content
                        or "Just a moment" in content
                    ):
                        print(f"[{url}] Still blocked — marking as blocked")
                        row["is_active"] = 0
                        row["status_reason"] = "Cloudflare blocked"
                        writer.writerow(row)
                        f_out.flush()
                        done_urls.add(url)
                        time.sleep(5)
                        continue

                if (
                    "This tutor is no longer available" in content
                    or "Page Not Found" in content
                    or res.status == 404
                ):
                    row["is_active"] = 0
                    row["status_reason"] = "Soft 404 / Removed"
                elif (
                    page.locator("h1").count() > 0
                    and row.get("Teacher_Name", "").lower()
                    in page.locator("h1").first.inner_text().lower()
                ):
                    row["is_active"] = 1
                    row["status_reason"] = "OK"
                else:
                    row["is_active"] = 0
                    row["status_reason"] = "No tutor name found"

                writer.writerow(row)
                f_out.flush()
                done_urls.add(url)
                print(f"[{url}] -> {row['status_reason']}")

                # Jittered rate limit
                time.sleep(random.uniform(0.5, 1.0))

            except Exception as e:
                row["is_active"] = 0
                row["status_reason"] = f"Error: {str(e)[:50]}"
                writer.writerow(row)
                f_out.flush()
                done_urls.add(url)
                print(f"[{url}] -> Error: {str(e)[:50]}")

    browser.close()
    print("\nDone. Results saved to scraped_verification_results.csv")
