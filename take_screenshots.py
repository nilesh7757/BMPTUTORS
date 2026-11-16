import os
from playwright.sync_api import sync_playwright

os.makedirs('presentation/assets/generated', exist_ok=True)

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page(viewport={'width': 1920, 'height': 1080})
    
    # 1. Landing page screenshot
    page.goto('http://localhost:8080', wait_until='networkidle')
    page.screenshot(path='presentation/assets/generated/screenshot_landing_directory.png')
    print('1. Captured landing page')
    
    # 2. Tutor modal screenshot (T1)
    page.evaluate('viewTutorDetails("T1")')
    page.wait_for_selector('#detail-modal:not(.hidden)')
    page.screenshot(path='presentation/assets/generated/screenshot_tutor_modal_t1.png')
    print('2. Captured tutor modal')
    page.evaluate('closeModal()')
    
    # 3. AI Fee Estimator screenshot
    page.evaluate('switchTab("fee-ai")')
    page.wait_for_selector('#tab-fee-ai:not(.hidden)')
    page.screenshot(path='presentation/assets/generated/screenshot_fee_estimator.png')
    print('3. Captured Fee Estimator')
    
    # 4. AI Bio Generator Studio screenshot
    page.evaluate('switchTab("bio-ai")')
    page.wait_for_selector('#tab-bio-ai:not(.hidden)')
    page.screenshot(path='presentation/assets/generated/screenshot_bio_generator.png')
    print('4. Captured Bio Generator')
    
    # 5. Visual Analytics screenshot
    page.evaluate('switchTab("analytics")')
    page.wait_for_selector('#tab-analytics:not(.hidden)')
    page.screenshot(path='presentation/assets/generated/screenshot_visual_analytics.png')
    print('5. Captured Visual Analytics')
    
    browser.close()
    print('All screenshots captured successfully!')
