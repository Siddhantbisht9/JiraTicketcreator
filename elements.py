from playwright.sync_api import sync_playwright
import keyboard  # pip install keyboard
import time

def track_jira_elements():
    with sync_playwright() as p:
        # Launch Edge (Chromium) — uses "msedge" channel
        browser = p.chromium.launch(channel="msedge", headless=False)
        page = browser.new_page()
        
        # Go to your Jira instance
        page.goto("https://aboutsib.atlassian.net")  # <- replace with your Jira URL
        print("Log in manually (or use existing Edge session), then press ENTER here to continue...")
        input()  # Wait for manual login or existing session detection
        
        print("Page loaded. Press 'q' in this terminal to scan visible elements.")
        
        # Wait until 'q' is pressed
        while True:
            if keyboard.is_pressed('q'):
                print("\n=== SCANNING VISIBLE ELEMENTS ===\n")
                
                # Scan all buttons
                all_buttons = page.locator("button").all()
                print("=== Visible Buttons ===")
                for i, btn in enumerate(all_buttons):
                    try:
                        text = btn.inner_text().strip()
                        print(f"{i}: '{text}' -> selector: button >> text='{text}'")
                    except:
                        pass
                
                # Scan all menu items
                all_menuitems = page.locator("[role='menuitem']").all()
                print("\n=== Visible Menu Items ===")
                for i, item in enumerate(all_menuitems):
                    try:
                        text = item.inner_text().strip()
                        print(f"{i}: '{text}' -> selector: [role='menuitem'] >> text='{text}'")
                    except:
                        pass
                
                print("\nScan complete. Press 'q' again to scan again, or Ctrl+C to exit.")
                
                # Avoid multiple triggers
                time.sleep(1)

if __name__ == "__main__":
    track_jira_elements()