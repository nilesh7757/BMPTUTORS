import asyncio
from playwright.async_api import async_playwright
import os

OUTPUT_DIR = "/home/nilesh7757/BMPTUTORS1/presentation/assets/playwright_deck"
os.makedirs(OUTPUT_DIR, exist_ok=True)

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            viewport={"width": 1600, "height": 1000},
            device_scale_factor=2
        )
        page = await context.new_page()

        # 1. Landing Page / Hero Banner
        print("Capturing 01_hero_and_stats...")
        await page.goto("http://localhost:8080", wait_until="networkidle")
        await page.wait_for_timeout(1000)
        hero = await page.query_selector("header") or await page.query_selector(".max-w-7xl")
        if hero:
            await hero.screenshot(path=f"{OUTPUT_DIR}/01_hero_and_stats.png")
        else:
            await page.screenshot(path=f"{OUTPUT_DIR}/01_hero_and_stats.png")

        # 2. Search & Filter Bar
        print("Capturing 02_search_and_filters...")
        filter_section = await page.query_selector("#filterSection") or await page.query_selector(".sticky")
        if filter_section:
            await filter_section.screenshot(path=f"{OUTPUT_DIR}/02_search_and_filters.png")

        # 3. Tutor Cards Directory View
        print("Capturing 03_tutor_cards_directory...")
        cards_container = await page.query_selector("#teachersList") or await page.query_selector("#resultsContainer")
        if cards_container:
            await cards_container.screenshot(path=f"{OUTPUT_DIR}/03_tutor_cards_directory.png")
        else:
            await page.screenshot(path=f"{OUTPUT_DIR}/03_tutor_cards_directory.png")

        # 4. Tutor Profile Modal (T1 with 5 offerings)
        print("Capturing 04_tutor_profile_modal...")
        await page.evaluate("if (typeof openModal === 'function') { openModal('T1'); }")
        await page.wait_for_timeout(1200)
        modal_content = await page.query_selector("#modalContent") or await page.query_selector(".relative.bg-white")
        if modal_content:
            await modal_content.screenshot(path=f"{OUTPUT_DIR}/04_tutor_profile_modal.png")
        else:
            await page.screenshot(path=f"{OUTPUT_DIR}/04_tutor_profile_modal.png")
        
        # Close modal
        await page.keyboard.press("Escape")
        await page.wait_for_timeout(500)

        # 5. Visual Analytics Tab - Map & Macro view
        print("Capturing 05_analytics_full_report...")
        await page.goto("http://localhost:8080/report.html", wait_until="networkidle")
        await page.wait_for_timeout(1500)
        await page.screenshot(path=f"{OUTPUT_DIR}/05_analytics_view_top.png")
        
        # Capture choropleth map container specifically
        map_el = await page.query_selector("#map") or await page.query_selector(".leaflet-container")
        if map_el:
            await map_el.screenshot(path=f"{OUTPUT_DIR}/05b_leaflet_choropleth_map.png")

        # 6. Scroll down to Venn Diagram & Subject Charts
        print("Capturing 06_venn_diagram_section...")
        venn_el = await page.query_selector("#vennContainer") or await page.query_selector("#vennStudio")
        if venn_el:
            await venn_el.screenshot(path=f"{OUTPUT_DIR}/06_venn_diagram_section.png")

        # 7. Demographics & Tables Section
        print("Capturing 07_analytics_tables_view...")
        await page.evaluate("window.scrollBy(0, 1000)")
        await page.wait_for_timeout(800)
        await page.screenshot(path=f"{OUTPUT_DIR}/07_analytics_tables_view.png")

        # 8. Fee Estimator Modal/Tool
        print("Capturing 08_fee_estimator_tool...")
        await page.goto("http://localhost:8080", wait_until="networkidle")
        await page.wait_for_timeout(600)
        await page.evaluate("""() => {
            const btn = Array.from(document.querySelectorAll('button')).find(b => b.innerText.includes('Fee Estimator'));
            if (btn) btn.click();
        }""")
        await page.wait_for_timeout(1000)
        modal = await page.query_selector("#feeEstimatorModal") or await page.query_selector(".fixed.inset-0:not(.hidden)")
        if modal:
            await modal.screenshot(path=f"{OUTPUT_DIR}/08_fee_estimator_tool.png")
        else:
            await page.screenshot(path=f"{OUTPUT_DIR}/08_fee_estimator_tool.png")

        # 9. AI Bio Generator Modal/Tool
        print("Capturing 09_ai_bio_generator_tool...")
        await page.keyboard.press("Escape")
        await page.wait_for_timeout(500)
        await page.evaluate("""() => {
            const btn = Array.from(document.querySelectorAll('button')).find(b => b.innerText.includes('Bio Generator'));
            if (btn) btn.click();
        }""")
        await page.wait_for_timeout(1000)
        bio_modal = await page.query_selector("#bioModal") or await page.query_selector(".fixed.inset-0:not(.hidden)")
        if bio_modal:
            await bio_modal.screenshot(path=f"{OUTPUT_DIR}/09_ai_bio_generator_tool.png")
        else:
            await page.screenshot(path=f"{OUTPUT_DIR}/09_ai_bio_generator_tool.png")

        # 10. Full page scrolling overview
        print("Capturing 10_full_directory_page...")
        await page.goto("http://localhost:8080", wait_until="networkidle")
        await page.wait_for_timeout(600)
        await page.screenshot(path=f"{OUTPUT_DIR}/10_full_directory_viewport.png")

        await browser.close()
        print(f"All assets successfully saved in {OUTPUT_DIR}")

asyncio.run(main())
