import csv
import os
import time
import random
from playwright.sync_api import sync_playwright

INPUT_CSV = "retry_unscraped.csv"
OUTPUT_CSV = "retry_unscraped_results.csv"

USER_AGENT = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)

# Load already-processed tutor IDs so we can resume after interruptions
done_ids = set()
if os.path.exists(OUTPUT_CSV) and os.path.getsize(OUTPUT_CSV) > 0:
    with open(OUTPUT_CSV, "r", encoding="utf-8") as f:
        for line in f:
            parts = line.strip().split(",")
            if parts and parts[0] != "tutor_id":
                done_ids.add(parts[0])

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
        writer = csv.writer(f_out)
        if write_header:
            writer.writerow(["tutor_id", "profile_url", "is_active", "status_reason"])

        for row in reader:
            t_id = row["tutor_id"]
            if t_id in done_ids:
                continue  # already verified

            url = row["profile_url"]
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
                        f"\n[{t_id}] Cloudflare blocked! Solve the challenge in the browser window, then press ENTER here to continue..."
                    )
                    input()
                    content = page.content()
                    if (
                        "Sorry, you have been blocked" in content
                        or "Attention Required!" in content
                        or "Just a moment" in content
                    ):
                        print(f"[{t_id}] Still blocked — marking as blocked")
                        writer.writerow([t_id, url, 0, "Cloudflare blocked"])
                        f_out.flush()
                        time.sleep(5)
                        continue

                if (
                    "This tutor is no longer available" in content
                    or "Page Not Found" in content
                    or res.status == 404
                ):
                    status = 0
                    reason = "Soft 404 / Removed"
                elif (
                    page.locator("h1").count() > 0
                    and row["original_name"].lower()
                    in page.locator("h1").first.inner_text().lower()
                ):
                    status = 1
                    reason = "OK"
                else:
                    status = 0
                    reason = "No tutor name found"

                writer.writerow([t_id, url, status, reason])
                f_out.flush()
                done_ids.add(t_id)
                print(f"[{t_id}] {url} -> {reason}")

                # Jittered rate limit to look less robotic
                time.sleep(random.uniform(0.5, 1.0))

            except Exception as e:
                writer.writerow([t_id, url, 0, f"Error: {str(e)[:50]}"])
                f_out.flush()
                print(f"[{t_id}] {url} -> Error: {str(e)[:50]}")

    browser.close()
