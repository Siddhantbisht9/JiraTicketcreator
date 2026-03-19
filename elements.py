from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(channel="msedge", headless=False)
    context = browser.new_context()
    page = context.new_page()

    page.goto("https://atlassian.net") #use your Jira url

    print("Login manually and navigate to the page with the Export button...")
    page.pause()  # Wait for you to do manual navigation

    # Find the three-dots button
    three_dots = page.get_by_test_id(
        "issue-navigator-action-export-issues.ui.filter-button--trigger"
    )
    print("Three-dots button found:", three_dots.count())

    # Click to open the dropdown
    three_dots.click()

    # Now find the actual Export menu item
    export_item = page.get_by_role("menuitem", name="Export Excel CSV (all fields)")
    export_item.wait_for(state="visible", timeout=10000)
    print("Export button is visible and ready to click.")

    # Optional: click it
    # export_item.click()

    browser.close()