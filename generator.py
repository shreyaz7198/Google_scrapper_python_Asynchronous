import sys
import os
import asyncio
import sqlite3
import aiohttp
from pathlib import Path
import pandas as pd

# Import Playwright Async API
from playwright.async_api import async_playwright

# ==========================================
# 🛠️ SECTION 1: STORAGE & INITIALIZATION
# ==========================================

def get_default_storage_path(file_name: str) -> Path:
    """Gets a safe, cross-platform default path in the Downloads folder."""
    downloads_dir = Path.home() / "Downloads"
    if not downloads_dir.exists():
        downloads_dir = Path.home()
    return downloads_dir / file_name

def init_excel_file(path: Path):
    """Creates a structured Excel file with headers if it doesn't exist."""
    if not path.exists():
        headers = ["Timestamp", "Keyword", "Place", "Name", "Contact Number", "Address", "Plus Code", "Location", "Website", "Timings"]
        df = pd.DataFrame(columns=headers)
        df.to_excel(path, index=False)
        print(f"✅ Created fresh Excel template at: {path}")

def init_sqlite_db(path: Path):
    """Creates a SQLite database and the leads table if it doesn't exist."""
    conn = sqlite3.connect(path)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS leads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
            keyword TEXT,
            place TEXT,
            name TEXT,
            contact_number TEXT,
            address TEXT,
            plus_code TEXT,
            location TEXT,
            website TEXT,
            timings TEXT
        )
    """)
    conn.commit()
    conn.close()
    print(f"✅ Initialized SQLite Database template at: {path}")

# ==========================================
# 🎛️ SECTION 2: USER INTERFACE & FLOW
# ==========================================

def display_google_sheets_instructions() -> str:
    """Shows clear step-by-step instructions for Apps Script Web App integration."""
    print("\n" + "="*60)
    print("📋 GOOGLE SHEETS SETUP INSTRUCTIONS FOR NON-TECH USERS")
    print("="*60)
    print("1. Open Google Sheets (Ensure you are logged into ONE Google account only).")
    print("2. Create a brand new Spreadsheet (or open an existing one).")
    print("3. In the top menu, go to: Extensions -> Apps Script.")
    print("4. Erase any code inside the editor window and paste the exact code below:\n")
    
    apps_script_code = """function doPost(e) {
  try {
    // Parse the incoming JSON data from Python
    var data = JSON.parse(e.postData.contents);
    
    // Open the active spreadsheet and get the first sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // AUTOMATED HEADER CHECK: If Row 1, Column 1 is empty, write headers first
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp", "Keyword", "Place", "Name", "Contact Number", 
        "Address", "Plus Code", "Location", "Website", "Timings"
      ]);
    }
    
    // Extract your specific data fields (with fallback text if missing)
    var timestamp   = new Date();
    var keyword     = data.keyword || "";
    var place       = data.place || "";
    var name        = data.name || "";
    var contact     = data.contact_number || "";
    var address     = data.address || "";
    var plusCode    = data.plus_code || "";
    var location    = data.location || "";
    var website     = data.website || "";
    var timings     = data.timings || "";
    
    // Append the data as a new row in your Sheet
    sheet.appendRow([timestamp, keyword, place, name, contact, address, plusCode, location, website, timings]);
    
    // Return success response to Python
    return ContentService.createTextOutput(JSON.stringify({"status": "success", "message": "Row added successfully"}))
                         .setMimeType(ContentService.MimeType.JSON);
                         
  } catch (error) {
    // Return error response if something goes wrong
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}"""
    print(apps_script_code)
    print("\n" + "-"*60)
    print("5. Click 'Save' (the floppy disk icon).")
    print("6. Click the blue 'Deploy' button -> Select 'New deployment'.")
    print("7. Click the Gear icon (Select type) -> Choose 'Web app'.")
    print("8. Change 'Who has access' to: 'Anyone'. (Crucial for the script to access it).")
    print("9. Click 'Deploy', authorize permissions if asked, and COPY the generated 'Web app URL'.")
    print("="*60 + "\n")
    
    web_app_url = ""
    while not web_app_url.startswith("http"):
        web_app_url = input("🔗 Paste your deployed Google Web App URL here to proceed: ").strip()
        if not web_app_url.startswith("http"):
            print("⚠️ Invalid URL. It should start with http:// or https://")
            
    return web_app_url

def get_search_criteria():
    """Validates and collects keywords and locations from the user."""
    print("\n--- Search Target Setup ---")
    
    while True:
        keywords_raw = input("🔍 Enter Search Keywords (comma-separated, e.g., Dentists, Cafes): ").strip()
        keywords = [k.strip() for k in keywords_raw.split(",") if k.strip()]
        if keywords:
            break
        print("⚠️ You must enter at least one keyword!")

    while True:
        places_raw = input("📍 Enter Target Places/Locations (comma-separated, e.g., London, New York): ").strip()
        places = [p.strip() for p in places_raw.split(",") if p.strip()]
        if places:
            break
        print("⚠️ You must enter at least one place!")
        
    return keywords, places

# ==========================================
# 💾 SECTION 3: ASYNC DATA STORAGE ROUTER
# ==========================================

async def save_scraped_data_async(config, payload, session=None):
    """Cleans up data-artifacts and stores payload using non-blocking calls where possible."""
    def clean_text(text):
        if not text:
            return ""
        for char in ["", "", "", "🌍", "📞", "🕒"]:
            text = text.replace(char, "")
        return text.strip()

    cleaned_payload = {
        "keyword": payload["keyword"],
        "place": payload["place"],
        "name": clean_text(payload["name"]),
        "contact_number": clean_text(payload["contact_number"]),
        "address": clean_text(payload["address"]),
        "plus_code": clean_text(payload["plus_code"]),
        "location": payload["location"],
        "website": payload["website"],
        "timings": clean_text(payload["timings"])
    }

    if config["type"] == "excel":
        # Filesystem operations are blocking by nature in standard Pandas; running in executor or keeping it quick
        file_path = config["path_or_url"]
        excel_payload = {
            "Timestamp": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Keyword": cleaned_payload["keyword"],
            "Place": cleaned_payload["place"],
            "Name": cleaned_payload["name"],
            "Contact Number": cleaned_payload["contact_number"],
            "Address": cleaned_payload["address"],
            "Plus Code": cleaned_payload["plus_code"],
            "Location": cleaned_payload["location"],
            "Website": cleaned_payload["website"],
            "Timings": cleaned_payload["timings"]
        }
        df_new = pd.DataFrame([excel_payload])
        if file_path.exists():
            try:
                df_old = pd.read_excel(file_path)
                headers = ["Timestamp", "Keyword", "Place", "Name", "Contact Number", "Address", "Plus Code", "Location", "Website", "Timings"]
                df_old = df_old.reindex(columns=headers)
                df_final = pd.concat([df_old, df_new], ignore_index=True)
            except Exception:
                df_final = df_new
        else:
            df_final = df_new
        df_final.to_excel(file_path, index=False)
        
    elif config["type"] == "sqlite":
        conn = sqlite3.connect(config["path_or_url"])
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO leads (keyword, place, name, contact_number, address, plus_code, location, website, timings)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (cleaned_payload["keyword"], cleaned_payload["place"], cleaned_payload["name"], cleaned_payload["contact_number"], 
              cleaned_payload["address"], cleaned_payload["plus_code"], cleaned_payload["location"], cleaned_payload["website"], cleaned_payload["timings"]))
        conn.commit()
        conn.close()
        
    elif config["type"] == "google_sheets" and session:
        try:
            async with session.post(config["path_or_url"], json=cleaned_payload) as response:
                if response.status != 200:
                    text = await response.text()
                    print(f"⚠️ Google Sheets warning response: {text}")
        except Exception as e:
            print(f"❌ Failed to broadcast to Google Sheets async: {e}")

# ==========================================
# 🎭 SECTION 4: SCRAPING ENGINES (PLAYWRIGHT ASYNC)
# ==========================================

async def detail_consumer_worker(worker_id, queue, context, config, session):
    """Parallel consumer worker that processes details from the URL queue."""
    page = await context.new_page()
    # Optimize network routing per worker context
    await page.route("**/*", lambda route: route.abort() if route.request.resource_type in ["image", "font", "media"] else route.continue_())
    
    while True:
        task = await queue.get()
        if task is None:
            # Poison pill received, shut down worker gracefully
            await page.close()
            queue.task_done()
            break
            
        url, keyword, place = task
        try:
            await page.goto(url, wait_until="domcontentloaded", timeout=25000)
            await page.wait_for_timeout(1000)
            
            name = await page.locator("h1").first.text_content() if await page.locator("h1").count() > 0 else "Unknown Name"
            address = await page.locator("button[data-item-id='address']").text_content() if await page.locator("button[data-item-id='address']").count() > 0 else ""
            contact = await page.locator("button[data-item-id*='phone:tel']").text_content() if await page.locator("button[data-item-id*='phone:tel']").count() > 0 else ""
            website = await page.locator("a[data-item-id='authority']").get_attribute("href") if await page.locator("a[data-item-id='authority']").count() > 0 else ""
            plus_code = await page.locator("button[data-item-id='oloc']").text_content() if await page.locator("button[data-item-id='oloc']").count() > 0 else ""
            timings = await page.locator("div[aria-label*='Hours']").first.get_attribute("aria-label") if await page.locator("div[aria-label*='Hours']").count() > 0 else ""
            
            payload = {
                "keyword": keyword, "place": place, "name": name, "contact_number": contact,
                "address": address, "plus_code": plus_code, "location": page.url, "website": website, "timings": timings
            }
            
            print(f"   ✨ [Worker {worker_id}] Scraped: {payload['name']}")
            await save_scraped_data_async(config, payload, session)
            
        except Exception:
            pass
        finally:
            queue.task_done()

async def run_async_playwright_scraper(config):
    print("\n🚀 Starting Phase 2: Playwright Headless Asynchronous Concurrent Engine...")
    
    # Initialize the centralized pipeline queue
    queue = asyncio.Queue(maxsize=50)
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        
        # Setup modern single session client for asynchronous API pushing
        async with aiohttp.ClientSession() as session:
            
            # Spin up a pool of concurrent consumer workers (Setting to 4 workers for high-speed stability)
            num_workers = 4
            workers = []
            for i in range(num_workers):
                worker = asyncio.create_task(detail_consumer_worker(i + 1, queue, context, config, session))
                workers.append(worker)
                
            feed_page = await context.new_page()
            await feed_page.route("**/*", lambda route: route.abort() if route.request.resource_type in ["image", "font", "media"] else route.continue_())
            
            for keyword in config["keywords"]:
                for place in config["places"]:
                    search_query = f"{keyword} {place}"
                    print(f"\n⚡ [Producer] Activating search layouts for: '{search_query}'")
                    
                    formatted_query = search_query.replace(" ", "+")
                    await feed_page.goto(f"https://www.google.com/maps/search/{formatted_query}")
                    await feed_page.wait_for_timeout(4000)
                    
                    left_panel_selector = "div[role='feed']"
                    if await feed_page.locator(left_panel_selector).count() == 0:
                        print("⚠️ Feed view container hidden. Skipping configuration block.")
                        continue
                        
                    panel = feed_page.locator(left_panel_selector)
                    processed_urls = set()
                    
                    scroll_attempts_without_new_data = 0
                    max_scrolls = 15
                    scrolls = 0
                    
                    while scrolls < max_scrolls:
                        current_locators = await feed_page.locator("a[href*='/maps/place/']").all()
                        current_urls = [await loc.get_attribute("href") for loc in current_locators if await loc.get_attribute("href")]
                        
                        new_urls = [url for url in current_urls if url not in processed_urls]
                        
                        if new_urls:
                            scroll_attempts_without_new_data = 0
                            print(f"🔄 [Producer] Found {len(new_urls)} new URLs. Queueing targets...")
                            for url in new_urls:
                                processed_urls.add(url)
                                # Feed URL task coordinates directly to our waiting concurrent consumer workers
                                await queue.put((url, keyword, place))
                        else:
                            scroll_attempts_without_new_data += 1
                        
                        # Command tracking view panel down
                        await panel.evaluate("element => element.scrollBy(0, element.scrollHeight)")
                        await feed_page.wait_for_timeout(2000)
                        
                        if scroll_attempts_without_new_data >= 4:
                            print("🛑 [Producer] Search feed markers exhausted.")
                            break
                            
                        scrolls += 1
                        
            # Wait until all queued items have finished processing
            await queue.join()
            
            # Dispatch poison pills to shut down consumer pool threads cleanly
            for _ in range(num_workers):
                await queue.put(None)
            await asyncio.gather(*workers)
            
        await browser.close()
    print("\n🏁 Phase 2 Async Processing Complete. System clean.")

# ==========================================
# 🚀 SECTION 5: INITIALIZATION FLOW
# ==========================================

def run_setup_flow():
    print("=========================================")
    print("🚀 Welcome to the Modern Lead Scraper CLI")
    print("=========================================")
    print("How would you like your data to be generated?")
    print("1. Excel file in your system")
    print("2. Google Sheet")
    print("3. Database file (.db)")
    print("Press any other key to Exit.")
    
    choice = input("Select an option: ").strip()
    config = {"type": None, "path_or_url": None, "keywords": [], "places": []}
    
    if choice == "1":
        config["type"] = "excel"
        config["path_or_url"] = get_default_storage_path("maps_scraped_leads.xlsx")
        init_excel_file(config["path_or_url"])
    elif choice == "3":
        config["type"] = "sqlite"
        config["path_or_url"] = get_default_storage_path("maps_scraped_leads.db")
        init_sqlite_db(config["path_or_url"])
    elif choice == "2":
        config["type"] = "google_sheets"
        config["path_or_url"] = display_google_sheets_instructions()
    else:
        print("❌ Exiting application. Goodbye!")
        sys.exit(0)
        
    config["keywords"], config["places"] = get_search_criteria()
    
    print("\n✅ Configuration Setup Ready!")
    return config

if __name__ == "__main__":
    session_config = run_setup_flow()
    
    # Bootstrap the asynchronous loop context for engine execution
    asyncio.run(run_async_playwright_scraper(session_config))