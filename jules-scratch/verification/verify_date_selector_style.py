from playwright.sync_api import sync_playwright, expect
import os

def run_verification():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        # Build the file path to the local HTML file
        file_path = "file://" + os.path.abspath("dist/index.html")
        page.goto(file_path)

        # Expect the year and month selectors to be visible
        year_select = page.locator("#yearSelect")
        month_select = page.locator("#monthSelect")

        expect(year_select).to_be_visible()
        expect(month_select).to_be_visible()

        # Check the default selected values (next month)
        # Based on the current date of 2025-10-28
        expect(year_select).to_have_value("2025")
        expect(month_select).to_have_value("11")

        # Take a screenshot
        screenshot_path = "jules-scratch/verification/verification_styled.png"
        page.screenshot(path=screenshot_path)

        print(f"Screenshot saved to {screenshot_path}")

        browser.close()

if __name__ == "__main__":
    run_verification()
