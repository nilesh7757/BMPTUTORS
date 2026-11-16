import asyncio
import re
from playwright.async_api import async_playwright

async def audit():
    console_errors = []
    console_warnings = []
    failed_requests = []
    audit_findings = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(viewport={"width": 1400, "height": 900})
        page = await context.new_page()

        page.on("console", lambda msg: 
            console_errors.append(f"[{msg.type}] {msg.text}") if msg.type == "error" 
            else console_warnings.append(f"[{msg.type}] {msg.text}") if msg.type == "warning" 
            else None
        )
        page.on("requestfailed", lambda req: 
            failed_requests.append(f"FAILED {req.method} {req.url} - {req.failure}")
        )
        page.on("response", lambda resp: 
            failed_requests.append(f"HTTP {resp.status} {resp.url}") if resp.status >= 400 
            else None
        )

        print("=== STEP 1: LOAD HOMEPAGE ===", flush=True)
        await page.goto("http://127.0.0.1:8080", wait_until="domcontentloaded")
        await page.wait_for_timeout(2000)

        # 1. Homepage KPIs
        stat_tutors = await page.inner_text("#stat-tutors")
        stat_bridge = await page.inner_text("#stat-bridge")
        stat_fee = await page.inner_text("#stat-fee")
        stat_active = await page.inner_text("#stat-active")

        print(f"Homepage KPIs: Tutors={stat_tutors}, Subjects={stat_bridge}, Rate={stat_fee}, Active={stat_active}", flush=True)
        if stat_tutors != "107,673":
            audit_findings.append(f"Homepage KPI error: stat-tutors is {stat_tutors}, expected 107,673")
        if stat_bridge != "403,221":
            audit_findings.append(f"Homepage KPI error: stat-bridge is {stat_bridge}, expected 403,221")
        if "864" not in stat_fee:
            audit_findings.append(f"Homepage KPI error: stat-fee is {stat_fee}, expected ₹864")
        if stat_active != "76,683":
            audit_findings.append(f"Homepage KPI error: stat-active is {stat_active}, expected 76,683")

        # 2. Directory Page Info
        page_info = await page.inner_text("#page-info")
        print(f"Directory Page Info: {page_info.strip()}", flush=True)
        if "107,673" not in page_info:
            audit_findings.append(f"Directory page-info mismatch: '{page_info}', expected '107,673'")

        # 3. Test Tutor Modal
        tutor_card = page.locator("#tutors-container > div").first
        if await tutor_card.count() > 0:
            await tutor_card.click()
            await page.wait_for_timeout(400)
            modal = page.locator("#tutor-modal")
            is_vis = await modal.is_visible()
            modal_name = await page.inner_text("#modal-name")
            print(f"Tutor modal open test: visible={is_vis}, name={modal_name}", flush=True)
            if not is_vis:
                audit_findings.append("Tutor profile modal did not open upon card click")
            await page.click("#btn-close-modal")
            await page.wait_for_timeout(200)

        print("\n=== STEP 2: AUDIT VISUAL ANALYTICS TAB ===", flush=True)
        await page.click("#tab-btn-analytics")
        await page.wait_for_timeout(1500)

        # Visual Analytics Subtitle
        analytics_sub = await page.inner_text("#analytics-scope-subtitle")
        print(f"Analytics Subtitle: {analytics_sub.strip()}", flush=True)
        if "107,673" not in analytics_sub or "403,221" not in analytics_sub:
            audit_findings.append(f"Visual Analytics subtitle mismatch: '{analytics_sub}'")

        # Figure 1 Subtitle & Note
        fig1_sub = await page.inner_text("#fig1-sample-subtitle")
        fig1_note = await page.locator("p:has-text('Values show count (N) and percentage share across all')").inner_text()
        print(f"Figure 1 Subtitle: {fig1_sub.strip()}", flush=True)
        print(f"Figure 1 Note: {fig1_note.strip()}", flush=True)
        if "403,221" not in fig1_sub:
            audit_findings.append(f"Figure 1 subtitle missing 403,221: '{fig1_sub}'")
        if "403,221" not in fig1_note or "106,502" not in fig1_note or "1,171" not in fig1_note:
            audit_findings.append(f"Figure 1 note math inconsistency: '{fig1_note}'")

        # Figure 2 Note
        fig2_note = await page.locator("p:has-text('Based on 105,840 tutors with experience')").inner_text()
        print(f"Figure 2 Note: {fig2_note.strip()}", flush=True)
        if "105,840" not in fig2_note or "1,268" not in fig2_note or "565" not in fig2_note:
            audit_findings.append(f"Figure 2 note math inconsistency: '{fig2_note}'")

        # Figure 3 (Venn Studio) Subtitle
        fig3_sub = await page.locator("#section-venn-studio p.text-slate-500").first.inner_text()
        print(f"Figure 3 Subtitle: {fig3_sub.strip()}", flush=True)
        if "107,673" not in fig3_sub:
            audit_findings.append(f"Figure 3 subtitle missing 107,673: '{fig3_sub}'")

        # Test Venn Calculation Tool
        print("Testing Venn Calculation tool...", flush=True)
        await page.select_option("#venn-select-1", "Mathematics")
        await page.select_option("#venn-select-2", "Physics")
        await page.select_option("#venn-select-3", "Chemistry")
        await page.wait_for_timeout(1000)
        venn_svg = page.locator("#venn-svg-wrapper svg")
        has_svg = await venn_svg.count() > 0
        triple_overlap = await page.inner_text("#venn-kpi-all3")
        print(f"Venn SVG rendered: {has_svg}, Triple overlap count: {triple_overlap}", flush=True)
        if not has_svg:
            audit_findings.append("Venn Diagram SVG failed to render")

        # Table 1 (Gender) Note
        tbl1_note = await page.locator("p:has-text('Total Sample Size N = 106,502 unique educators')").inner_text()
        print(f"Table 1 Note: {tbl1_note.strip()}", flush=True)
        if "106,502" not in tbl1_note or "62,440" not in tbl1_note or "43,947" not in tbl1_note:
            audit_findings.append(f"Table 1 note math inconsistency: '{tbl1_note}'")

        # Table 2 (Fees) Note
        tbl2_note = await page.locator("p:has-text('Hourly fees benchmarked strictly within the realistic tuition band')").inner_text()
        print(f"Table 2 Note: {tbl2_note.strip()}", flush=True)
        if "78,757" not in tbl2_note or "28,916" not in tbl2_note:
            audit_findings.append(f"Table 2 note math inconsistency: '{tbl2_note}'")

        # Table 3 (Languages) Badge
        tbl3_badge = await page.inner_text("#lang-sample-badge")
        print(f"Table 3 Language Badge: {tbl3_badge.strip()}", flush=True)
        if "106,405" not in tbl3_badge:
            audit_findings.append(f"Table 3 badge missing 106,405: '{tbl3_badge}'")

        # Scope Switcher Test
        print("Testing Scope Switcher to 'Verified Active'...", flush=True)
        btn_active = page.locator("#btn-scope-active")
        await btn_active.click()
        await page.wait_for_timeout(800)
        active_sub = await page.inner_text("#analytics-scope-subtitle")
        act_stat_tutors = await page.inner_text("#stat-tutors")
        act_stat_subjects = await page.inner_text("#stat-bridge")
        print(f"Active Scope: subtitle='{active_sub.strip()}', tutors={act_stat_tutors}, subjects={act_stat_subjects}", flush=True)
        if "76,683" not in active_sub or "308,162" not in active_sub:
            audit_findings.append(f"Active scope subtitle mismatch: '{active_sub}'")
        if act_stat_tutors != "76,683" or act_stat_subjects != "308,162":
            audit_findings.append(f"Active scope KPI mismatch: {act_stat_tutors}, {act_stat_subjects}")

        # Switch back to All Master
        await page.click("#btn-scope-all")
        await page.wait_for_timeout(500)

        print("\n=== STEP 3: AUDIT AI FEE ESTIMATOR TOOL ===", flush=True)
        await page.click("#tab-btn-fee-ai")
        await page.wait_for_timeout(800)

        # Default Fee Math (15 Hours)
        pred_fee_def = await page.inner_text("#ai-pred-fee")
        earn_def = await page.inner_text("#ai-monthly-earning")
        approx_def = await page.inner_text("#ai-monthly-hours-approx")
        print(f"Fee Estimator (Default 15h): Rate={pred_fee_def}, Monthly={earn_def}, Approx={approx_def}", flush=True)

        rate_def_num = int(re.sub(r"[^\d]", "", pred_fee_def))
        earn_def_num = int(re.sub(r"[^\d]", "", earn_def))
        approx_def_num = int(re.sub(r"[^\d]", "", approx_def))
        if earn_def_num != rate_def_num * approx_def_num:
            audit_findings.append(f"Default Fee Math mismatch: {rate_def_num} * {approx_def_num} = {rate_def_num * approx_def_num}, UI={earn_def_num}")

        # Move Hours Slider to 16 and trigger calculation
        await page.evaluate("""async () => {
            const slider = document.getElementById('ai-hours-slider');
            slider.value = '16';
            onFeeParamChange();
            await calculateFeePrediction();
        }""")
        await page.wait_for_timeout(800)

        pred_fee_16 = await page.inner_text("#ai-pred-fee")
        earn_16 = await page.inner_text("#ai-monthly-earning")
        approx_16 = await page.inner_text("#ai-monthly-hours-approx")
        print(f"Fee Estimator (16 Hours/Week): Rate={pred_fee_16}, Monthly={earn_16}, Approx={approx_16}", flush=True)

        rate_num = int(re.sub(r"[^\d]", "", pred_fee_16))
        earn_num = int(re.sub(r"[^\d]", "", earn_16))
        approx_num = int(re.sub(r"[^\d]", "", approx_16))

        expected_math = rate_num * approx_num
        print(f"Fee Calculation Verify: {rate_num} * {approx_num} = {expected_math} (UI displays {earn_num})", flush=True)
        if earn_num != expected_math:
            audit_findings.append(f"Fee Estimator math mismatch: {rate_num} * {approx_num} = {expected_math}, UI={earn_num}")

        print("\n=== STEP 4: AUDIT AI BIO GENERATOR TOOL (Bio Studio) ===", flush=True)
        await page.click("#tab-btn-bio-ai")
        await page.wait_for_timeout(500)
        await page.fill("#bio-input-name", "Dr. Rajesh Sharma")
        await page.fill("#bio-input-subj", "Mathematics & Physics")
        await page.click("#btn-generate-bio")
        await page.wait_for_timeout(1200)
        bio_text = await page.input_value("#bio-output-textarea")
        print(f"Bio Studio Output: {len(bio_text.strip())} characters generated", flush=True)
        if len(bio_text.strip()) < 50:
            audit_findings.append("Bio Studio failed to generate profile bio")

        print("\n=== STEP 5: AUDIT ABOUT PROJECT TAB ===", flush=True)
        await page.click("#tab-btn-about")
        await page.wait_for_timeout(500)
        about_text = await page.inner_text("#tab-about")
        if "107,723" in about_text:
            audit_findings.append("About page still contains raw number 107,723")
        if "327,720" in about_text:
            audit_findings.append("About page still contains raw number 327,720")
        print(f"About page contains clean '107,673': {'107,673' in about_text}", flush=True)

        print("\n=== STEP 6: AUDIT REPORT PAGE (/report) ===", flush=True)
        await page.goto("http://127.0.0.1:8080/report", wait_until="domcontentloaded")
        await page.wait_for_timeout(1000)
        report_text = await page.inner_text("body")
        if "107,723" in report_text:
            audit_findings.append("Report page contains raw number 107,723")
        if "107,673" not in report_text:
            audit_findings.append("Report page missing clean number 107,673")
        print(f"Report page contains clean '107,673': {'107,673' in report_text}", flush=True)

        await browser.close()

    print("\n" + "="*50, flush=True)
    print("PLAYWRIGHT AUDIT COMPLETE", flush=True)
    print(f"Total Console Errors: {len(console_errors)}", flush=True)
    for err in console_errors:
        print(f"  ❌ {err}", flush=True)
    print(f"Total Console Warnings: {len(console_warnings)}", flush=True)
    for warn in console_warnings:
        print(f"  ⚠️ {warn}", flush=True)
    print(f"Total Failed Requests: {len(failed_requests)}", flush=True)
    for freq in failed_requests:
        print(f"  🚫 {freq}", flush=True)
    print(f"Total Inconsistencies / Mistakes Found: {len(audit_findings)}", flush=True)
    for f in audit_findings:
        print(f"  🚨 {f}", flush=True)
    print("="*50, flush=True)

if __name__ == "__main__":
    asyncio.run(audit())
