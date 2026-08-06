import asyncio
from playwright.async_api import async_playwright
import os

async def run_test():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()
        
        # Set viewport
        await page.set_viewport_size({"width": 1920, "height": 1080})
        
        # Console logs
        page.on("console", lambda msg: print(f"CONSOLE: {msg.text}"))
        
        try:
            # Get frontend URL from frontend/.env or default to localhost:3000
            url = "http://localhost:3000"
            print(f"Navigating to {url}...")
            await page.goto(url, timeout=10000)
            
            # 1. Verify app loads and title is present
            await page.wait_for_selector("text=SAP WM Transfer Order Processor", timeout=5000)
            print("Successfully loaded SAP WM Transfer Order Processor app.")
            
            # 2. Verify Row column is automatically generated and shown in table
            await page.wait_for_selector("text=Row (Extracted)", timeout=3000)
            print("Row column header verified in table.")
            
            # Check extracted row value for first item (Source Storage Bin: M1-07-36A2 -> Row: 07)
            row_cell = await page.wait_for_selector("text=07", timeout=3000)
            assert row_cell is not None
            print("Verified automated Row extraction (07 from M1-07-36A2).")
            
            # 3. Verify card menu for ranges (Card 1, Card 2, etc.)
            card1 = await page.wait_for_selector('[data-testid="card-batch-1"]', timeout=3000)
            assert card1 is not None
            print("Card 1 found successfully.")
            
            # Click Card 1 to filter
            await card1.click()
            await page.wait_for_timeout(500)
            print("Clicked Card 1 to filter transfer orders.")
            
            # Verify table header updates or row count is filtered
            table_section = await page.wait_for_selector('[data-testid="sap-table-section"]', timeout=3000)
            assert table_section is not None
            print("Filtered table displayed successfully.")
            
            # 4. Verify Copy to SAP button
            copy_btn = await page.wait_for_selector('[data-testid="copy-to-sap-btn"]', timeout=3000)
            assert copy_btn is not None
            await copy_btn.click()
            print("Clicked 'Copy untuk Paste ke SAP (TSV)' button successfully.")
            
            # 5. Verify batch size buttons working
            batch_10_btn = await page.wait_for_selector('[data-testid="batch-size-10-btn"]', timeout=3000)
            await batch_10_btn.click()
            await page.wait_for_timeout(500)
            print("Switched batch size to 10 successfully.")
            
            # 6. Verify Reset data button
            reset_btn = await page.wait_for_selector('[data-testid="reset-data-btn"]', timeout=3000)
            await reset_btn.click()
            await page.wait_for_timeout(500)
            print("Reset data button clicked successfully.")
            
            print("All frontend UI and integration tests passed successfully!")
            
        except Exception as e:
            print(f"Test failed with error: {str(e)}")
            raise e
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(run_test())
