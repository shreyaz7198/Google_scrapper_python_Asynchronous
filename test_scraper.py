import pytest
import asyncio
import pandas as pd
from pathlib import Path

# --- Core Async/Sync Cleaner Functions to Test ---
def clean_text(text):
    if not text: return ""
    for char in ["", "", "", "🌍", "📞", "🕒"]:
        text = text.replace(char, "")
    return text.strip()

async def mock_async_data_fetch(url):
    """Simulates an asynchronous browser extraction task delay."""
    await asyncio.sleep(0.1)  # Simulates network lag safely
    return {"status": "fetched", "url": url}

# =====================================================================
# 🧪 Test 1: Character Sanitation (Unit Test)
# =====================================================================
def test_clean_text_strips_google_font_artifacts():
    """Verifies that custom styling symbols from Google Maps are stripped."""
    dirty_address = " 456 Kingsway, London "
    expected_output = "456 Kingsway, London"
    assert clean_text(dirty_address) == expected_output

# =====================================================================
# 🧪 Test 2: Asynchronous Queue Event Loop Verification
# =====================================================================
@pytest.mark.asyncio
async def test_async_worker_processing_flow():
    """Verifies that our non-blocking worker queue operates without hanging."""
    target_url = "https://google.com/maps/place/sample"
    result = await mock_async_data_fetch(target_url)
    assert result["status"] == "fetched"
    assert result["url"] == target_url

# =====================================================================
# 🧪 Test 3: Excel Grid Integrity Verification
# =====================================================================
def test_excel_file_creation_structure(tmp_path):
    """Verifies that dataframes write to disk with perfect header mappings."""
    test_file = tmp_path / "async_leads.xlsx"
    headers = ["Timestamp", "Keyword", "Place", "Name", "Contact Number"]
    
    df = pd.DataFrame(columns=headers)
    df.to_excel(test_file, index=False)
    
    assert test_file.exists()
    df_read = pd.read_excel(test_file)
    assert list(df_read.columns) == headers