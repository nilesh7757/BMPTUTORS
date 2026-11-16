import asyncio
from playwright.async_api import async_playwright
import os

OUTPUT_DIR = "/home/nilesh7757/BMPTUTORS1/presentation/assets/playwright_deck"
os.makedirs(OUTPUT_DIR, exist_ok=True)

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        # 1920x1080 with 2x pixel ratio for ultra-crisp presentation visuals
        context = await browser.new_context(
            viewport={"width": 1920, "height": 1080},
            device_scale_factor=2
        )
        page = await context.new_page()

        # 1. Landing Page / Hero Banner
        print("Capturing 01_hero_and_stats...")
        await page.goto("http://localhost:8080", wait_until="networkidle")
        await page.wait_for_timeout(1000)
        # Capture hero + stats section
        await page.screenshot(path=f"{OUTPUT_DIR}/01_hero_and_stats.png", clip={"x": 0, "y": 0, "width": 1920, "height": 650})

        # 2. Filter Bar & Search Utility
        print("Capturing 02_search_and_filters...")
        await page.evaluate("window.scrollTo(0, 320)")
        await page.wait_for_timeout(500)
        await page.screenshot(path=f"{OUTPUT_DIR}/02_search_and_filters.png", clip={"x": 0, "y": 320, "width": 1920, "height": 550})

        # 3. Tutor Cards Directory View
        print("Capturing 03_tutor_cards_directory...")
        await page.evaluate("window.scrollTo(0, 520)")
        await page.wait_for_timeout(800)
        await page.screenshot(path=f"{OUTPUT_DIR}/03_tutor_cards_directory.png", clip={"x": 0, "y": 520, "width": 1920, "height": 900})

        # 4. Tutor Profile Modal (T1 with 5 offerings)
        print("Capturing 04_tutor_profile_modal...")
        await page.evaluate("if (typeof openModal === 'function') { openModal('T1'); }")
        await page.wait_for_timeout(1200)
        await page.screenshot(path=f"{OUTPUT_DIR}/04_tutor_profile_modal.png")

        # Close modal
        await page.keyboard.press("Escape")
        await page.wait_for_timeout(500)

        # 5. Visual Analytics Tab - Interactive Map
        print("Capturing 05_analytics_choropleth_map...")
        # Click on Visual Analytics tab
        analytics_tab = await page.query_selector("button:has-text('Visual Analytics'), [onclick*='analytics'], a[href*='analytics']")
        if analytics_tab:
            await analytics_tab.click()
            await page.wait_for_timeout(1500)
        else:
            await page.goto("http://localhost:8080/report.html", wait_until="networkidle")
            await page.wait_for_timeout(1500)
        
        await page.screenshot(path=f"{OUTPUT_DIR}/05_analytics_choropleth_map.png", clip={"x": 0, "y": 0, "width": 1920, "height": 950})

        # 6. Visual Analytics - Venn Diagram Studio & Subjects
        print("Capturing 06_analytics_venn_and_subjects...")
        await page.evaluate("window.scrollTo(0, 950)")
        await page.wait_for_timeout(1000)
        await page.screenshot(path=f"{OUTPUT_DIR}/06_analytics_venn_and_subjects.png", clip={"x": 0, "y": 950, "width": 1920, "height": 900})

        # 7. Visual Analytics - Experience & Gender Demographics
        print("Capturing 07_analytics_tables_and_demographics...")
        await page.evaluate("window.scrollTo(0, 1850)")
        await page.wait_for_timeout(1000)
        await page.screenshot(path=f"{OUTPUT_DIR}/07_analytics_tables_and_demographics.png", clip={"x": 0, "y": 1850, "width": 1920, "height": 950})

        # 8. Fee Estimator Tool
        print("Capturing 08_fee_estimator_tool...")
        await page.goto("http://localhost:8080", wait_until="networkidle")
        fee_tab = await page.query_selector("button:has-text('Fee Estimator'), [onclick*='fee'], a[href*='fee']")
        if fee_tab:
            await fee_tab.click()
            await page.wait_for_timeout(1000)
        await page.screenshot(path=f"{OUTPUT_DIR}/08_fee_estimator_tool.png")

        # 9. AI Bio Generator Tool
        print("Capturing 09_ai_bio_generator_tool...")
        bio_tab = await page.query_selector("button:has-text('Bio Generator'), [onclick*='bio'], a[href*='bio']")
        if bio_tab:
            await bio_tab.click()
            await page.wait_for_timeout(1000)
        await page.screenshot(path=f"{OUTPUT_DIR}/09_ai_bio_generator_tool.png")

        await browser.close()
        print(f"All assets successfully saved in {OUTPUT_DIR}")

asyncio.run(main())
