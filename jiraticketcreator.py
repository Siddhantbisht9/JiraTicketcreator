from playwright.sync_api import sync_playwright
import pandas as pd
import time
import sys

CSV_FILE = r"Path/JiraTicketcreator/JiraTicket.xlsx" #Add the ticket summary and description here
JIRA_URL = "https://.atlassian.net" #your Jira url
USER_DATA_DIR = r"path/to/your/edge_profile"    
WAIT = 15000  

def wait_for_create_button(page):
    return page.wait_for_selector(
        "button[data-testid='atlassian-navigation--create-button']",
        timeout=WAIT
    )

def maybe_handle_login(page):
    """
    If we aren't yet on a Jira page with the top bar, let the user complete SSO/login
    once in the visible Edge, then continue.
    """
    url = page.url.lower()
    if ("login" in url) or ("microsoftonline.com" in url) or ("id.atlassian.com" in url):
        print("\nLooks like a login/SSO page is shown in Edge.")
        print("Please complete the login in the visible Edge window.")
        print("When you see Jira fully loaded (top bar visible), press ENTER here to continue...")
        try:
            input()
        except KeyboardInterrupt:
            sys.exit(1)
        wait_for_create_button(page)
def fill_summary(page, title):
    for selector in [
        "[data-testid='issue-field-summary'] input",
        "#summary-field",
        "input[name='summary']",
        "input#summary",
    ]:
        try:
            page.wait_for_selector(selector, timeout=WAIT)
            page.fill(selector, "")
            page.fill(selector, title)
            return True
        except Exception:
            continue
    return False
def fill_description(page, desc):
    try:
        page.fill("textarea#description-field", desc)
        return True
    except Exception:
        pass
    for sel in [
        "[data-testid='issue-field-description'] [contenteditable='true']",
        "div[contenteditable='true'][role='textbox']",
        "div[contenteditable='true'][aria-label='Main content area']",
    ]:
        try:
            page.wait_for_selector(sel, timeout=3000)
            try:
                page.locator(sel).fill(desc)
            except Exception:
                page.locator(sel).click()
                page.keyboard.insert_text(desc)
            return True
        except Exception:
            continue
    for frame in page.frames:
        try:
            el = frame.wait_for_selector("body[contenteditable], div[contenteditable='true']", timeout=1500)
            el.click()
            frame.keyboard.insert_text(desc)
            return True
        except Exception:
            continue
    return False
def run():
    df = pd.read_csv(CSV_FILE)
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            channel="msedge",        
            headless=False,          
            viewport={"width": 1400, "height": 900}
        )
        context.set_default_timeout(WAIT)
        context.set_default_navigation_timeout(60000)  
        page = context.new_page()
        page.goto(JIRA_URL, wait_until="domcontentloaded", timeout=90000)
        try:
            wait_for_create_button(page)
        except Exception:
            maybe_handle_login(page)
        for _, row in df.iterrows():
            title = str(row["Title"]).strip()
            desc = str(row["Description"]).strip()
            print(f"Creating: {title}")
            try:
                page.click("button[data-testid='atlassian-navigation--create-button']", timeout=WAIT)
            except Exception:
                page.goto(f"{JIRA_URL.rstrip('/')}/secure/CreateIssue!default.jspa",
                          wait_until="domcontentloaded", timeout=90000)
            try:
                page.wait_for_selector(
                    "[data-testid='issue-field-summary'] input, #summary-field, input[name='summary'], input#summary",
                    timeout=WAIT
                )
            except Exception:
                page.screenshot(path="create_form_missing.png", full_page=True)
                with open("create_form_missing.html", "w", encoding="utf-8") as f:
                    f.write(page.content())
                print("Create form not visible; saved create_form_missing.* for troubleshooting.")
                continue
            if not fill_summary(page, title):
                print("Could not fill Summary; skipping this row.")
                continue
            if not fill_description(page, desc):
                print("Could not fill Description; proceeding with Summary only.")
            created = False
            for sel in [
                "button:has-text('Create')",
                "button[type='submit'][data-testid*='create-button']",
                "button[type='submit']",
            ]:
                try:
                    page.click(sel, timeout=WAIT)
                    created = True
                    break
                except Exception:
                    continue
            if not created:
                page.screenshot(path="create_button_missing.png", full_page=True)
                with open("create_button_missing.html", "w", encoding="utf-8") as f:
                    f.write(page.content())
                print("Create submit not clicked; saved create_button_missing.* for troubleshooting.")
                continue
            page.wait_for_timeout(2000)
        print("DONE. All rows processed.")
if __name__ == "__main__":
    run()


    