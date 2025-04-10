import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
import random
import sys
import re
from datetime import datetime
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys 
import argparse
import subprocess
try:
    from gmail_api import GmailAPI  # Import the GmailAPI class
    print("GmailAPI module imported successfully")
except ImportError as e:
    print("GmailAPI module not found. Error:", str(e))
    print("Please make sure you have the gmail_api.py file in your project directory")
    GmailAPI = None

class WeatherLink:
    def __init__(self, url, debug=False, export_email="teamdavcast@gmail.com", init_browser=True):
        self.url = url
        self.debug = debug
        self.export_email = export_email
        print(f"Initializing WeatherLink scraper with debug mode: {'ON' if debug else 'OFF'}")
        print(f"Export email: {export_email}")
        self.use_api = True  # Default to True, can be set to False later
        
        # Only initialize browser if requested
        if init_browser:
            self.driver = self.setup_browser()
        else:
            self.driver = None

    def setup_browser(self):
        """Set up the WebDriver."""
        try:
            print("Setting up Chrome browser...")
            options = Options()
            # Only use headless mode if not in debug mode
            if not self.debug:
                # Uncomment the line below if you want headless mode (no visible browser)
                # options.add_argument("--headless")
                pass
            options.add_argument("--disable-gpu")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-dev-shm-usage")
            
            # Add window size for better visibility
            options.add_argument("--window-size=1920,1080")
            
            # Try direct approach with Chrome first
            try:
                print("Installing/locating Chrome driver...")
                service = Service(ChromeDriverManager().install())
                
                print("Starting Chrome browser...")
                driver = webdriver.Chrome(service=service, options=options)
            except Exception as e1:
                print(f"Error with automatic ChromeDriver: {e1}")
                print("Trying alternative approach...")
                
                # Alternative method: Try to find Chrome in common locations
                chrome_paths = [
                    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
                    "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
                    os.environ.get("LOCALAPPDATA", "") + "\\Google\\Chrome\\Application\\chrome.exe"
                ]
                
                chrome_path = None
                for path in chrome_paths:
                    if os.path.exists(path):
                        chrome_path = path
                        print(f"Found Chrome at: {chrome_path}")
                        break
                        
                if chrome_path:
                    # Try to determine Chrome version
                    try:
                        import re
                        from subprocess import check_output
                        # Fix the backslash issue by preparing the string separately
                        escaped_path = chrome_path.replace("\\", "\\\\")
                        cmd = 'wmic datafile where name="' + escaped_path + '" get Version /value'
                        version_output = check_output(cmd, shell=True)
                        version_match = re.search(b'Version=([\\d\\.]+)', version_output)
                        if version_match:
                            chrome_version = version_match.group(1).decode('utf-8')
                            print(f"Chrome version: {chrome_version}")
                            
                            # Check if we already have a local chromedriver.exe
                            if os.path.exists("chromedriver.exe"):
                                print("Using existing chromedriver.exe from current directory")
                                service = Service("chromedriver.exe")
                                driver = webdriver.Chrome(service=service, options=options)
                                return driver
                            
                            # Get the major version
                            major_version = chrome_version.split('.')[0]
                            major_version_num = int(major_version)
                            
                            # Import required modules
                            import urllib.request
                            import zipfile
                            import tempfile
                            import json
                            import shutil
                            
                            try:
                                # For Chrome 115+, use Chrome for Testing API
                                if major_version_num >= 115:
                                    print(f"Using Chrome for Testing API for Chrome {major_version}+")
                                    
                                    # Create a temporary directory for the download
                                    with tempfile.TemporaryDirectory() as tmpdirname:
                                        # Get list of available versions
                                        print("Fetching available ChromeDriver versions...")
                                        versions_url = "https://googlechromelabs.github.io/chrome-for-testing/latest-versions-per-milestone.json"
                                        with urllib.request.urlopen(versions_url) as response:
                                            versions_data = json.loads(response.read().decode('utf-8'))
                                        
                                        # Find matching version for the major version
                                        if major_version in versions_data.get('milestones', {}):
                                            chromedriver_version = versions_data['milestones'][major_version]['version']
                                            print(f"Found matching ChromeDriver version: {chromedriver_version}")
                                            
                                            # Determine the platform
                                            print("Determining platform...")
                                            if os.name == 'nt':  # Windows
                                                # Check if system is 32-bit or 64-bit
                                                import platform
                                                if platform.machine().endswith('64'):
                                                    platform_name = "win64"
                                                else:
                                                    platform_name = "win32"
                                            else:
                                                platform_name = "linux64"  # Default to Linux 64-bit
                                            
                                            print(f"Detected platform: {platform_name}")
                                            
                                            # Construct download URL
                                            download_url = f"https://storage.googleapis.com/chrome-for-testing-public/{chromedriver_version}/{platform_name}/chromedriver-{platform_name}.zip"
                                            print(f"Downloading ChromeDriver from: {download_url}")
                                            
                                            # Download ChromeDriver
                                            zip_path = os.path.join(tmpdirname, "chromedriver.zip")
                                            try:
                                                urllib.request.urlretrieve(download_url, zip_path)
                                                
                                                # Extract the zip file
                                                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                                                    zip_ref.extractall(tmpdirname)
                                                
                                                # Find the chromedriver executable in the extracted files
                                                chromedriver_path = None
                                                for root, dirs, files in os.walk(tmpdirname):
                                                    for file in files:
                                                        if file == "chromedriver.exe":
                                                            chromedriver_path = os.path.join(root, file)
                                                            break
                                                    if chromedriver_path:
                                                        break
                                                
                                                if chromedriver_path and os.path.exists(chromedriver_path):
                                                    print(f"Found ChromeDriver executable at: {chromedriver_path}")
                                                    # Copy to current directory as backup
                                                    shutil.copy2(chromedriver_path, "chromedriver.exe")
                                                    print("Copied ChromeDriver to current directory")
                                                    
                                                    # Use the found ChromeDriver
                                                    service = Service(chromedriver_path)
                                                    driver = webdriver.Chrome(service=service, options=options)
                                                else:
                                                    print("ChromeDriver executable not found in the extracted files")
                                                    print("Falling back to manually downloaded ChromeDriver")
                                                    if os.path.exists("chromedriver.exe"):
                                                        service = Service("chromedriver.exe")
                                                        driver = webdriver.Chrome(service=service, options=options)
                                                    else:
                                                        raise Exception("ChromeDriver not found. Please download manually.")
                                            except Exception as download_error:
                                                print(f"Error during ChromeDriver download/extraction: {download_error}")
                                                print("Falling back to manually downloaded ChromeDriver")
                                                if os.path.exists("chromedriver.exe"):
                                                    service = Service("chromedriver.exe")
                                                    driver = webdriver.Chrome(service=service, options=options)
                                                else:
                                                    raise Exception("ChromeDriver not found. Please download manually.")
                                        else:
                                            print(f"No matching ChromeDriver version found for Chrome {major_version}")
                                            print("Falling back to manually downloaded ChromeDriver")
                                            if os.path.exists("chromedriver.exe"):
                                                service = Service("chromedriver.exe")
                                                driver = webdriver.Chrome(service=service, options=options)
                                            else:
                                                raise Exception("ChromeDriver not found. Please download manually.")
                                else:
                                    # For older Chrome versions (pre-115), use the old method
                                    print(f"Using legacy method for Chrome {major_version}")
                                    download_url = f"https://chromedriver.storage.googleapis.com/LATEST_RELEASE_{major_version}"
                                    
                                    # Get the exact version matching the Chrome version
                                    with urllib.request.urlopen(download_url) as response:
                                        exact_version = response.read().decode('utf-8')
                                    
                                    # Download the ChromeDriver zip file
                                    print(f"Downloading ChromeDriver version {exact_version}...")
                                    driver_url = f"https://chromedriver.storage.googleapis.com/{exact_version}/chromedriver_win32.zip"
                                    
                                    # Create a temporary directory for the download
                                    with tempfile.TemporaryDirectory() as tmpdirname:
                                        zip_path = os.path.join(tmpdirname, "chromedriver.zip")
                                        urllib.request.urlretrieve(driver_url, zip_path)
                                        
                                        # Extract the zip file
                                        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                                            zip_ref.extractall(tmpdirname)
                                        
                                        # Find and use the chromedriver executable
                                        chromedriver_path = os.path.join(tmpdirname, "chromedriver.exe")
                                        if os.path.exists(chromedriver_path):
                                            print(f"Using ChromeDriver from: {chromedriver_path}")
                                            service = Service(chromedriver_path)
                                            driver = webdriver.Chrome(service=service, options=options)
                                        else:
                                            raise Exception("ChromeDriver not found in the extracted files")
                            except Exception as download_err:
                                print(f"Error downloading ChromeDriver: {download_err}")
                                # Try to use local chromedriver.exe
                                if os.path.exists("chromedriver.exe"):
                                    print("Falling back to existing chromedriver.exe")
                                    service = Service("chromedriver.exe")
                                    driver = webdriver.Chrome(service=service, options=options)
                                else:
                                    raise
                    except Exception as e2:
                        print(f"Error determining Chrome version: {e2}")
                        print("Trying with default ChromeDriver...")
                        
                        # Last resort: Use direct ChromeDriver path
                        chromedriver_path = "./chromedriver.exe"
                        if os.path.exists(chromedriver_path):
                            service = Service(chromedriver_path)
                            driver = webdriver.Chrome(service=service, options=options)
                        else:
                            raise Exception("ChromeDriver not found. Please download it from https://chromedriver.chromium.org/downloads")
                else:
                    print("Chrome not found in common locations")
                    print("Please install Chrome or specify its location")
                    raise Exception("Chrome browser not found")
            
            print(f"Navigating to URL: {self.url}")
            driver.get(self.url)
            print("Chrome browser started successfully")
            return driver
        except Exception as e:
            print(f"ERROR setting up browser: {e}")
            print("Please make sure Chrome is installed on your system")
            print("You may need to download ChromeDriver manually:")
            print("1. Find your Chrome version in 'Help > About Google Chrome'")
            print("2. Download matching ChromeDriver from: https://chromedriver.chromium.org/downloads")
            print("3. Place chromedriver.exe in the same folder as this script")
            raise

    def save_data(self, df, filename="weather_data.csv"):
        """Save the extracted data to a CSV file."""
        if df is not None:
            df.to_csv(filename, index=False)
            print(f"Data saved to {filename}")
        else:
            print("No data to save.")
    
    def set_date_to_current(self):
        """Find the date selector and set it to the current system date."""
        try:
            print("Attempting to set date to current system date...")
            self.take_screenshot("before_date_selection")
            
            # Get current system date
            current_date = datetime.now()
            current_month = current_date.month
            current_day = current_date.day
            current_year = current_date.year
            month_names = ['January', 'February', 'March', 'April', 'May', 'June', 
                           'July', 'August', 'September', 'October', 'November', 'December']
            current_month_name = month_names[current_month - 1]
            
            print(f"Current system date: {current_month}/{current_day}/{current_year} ({current_month_name})")
            
            # Enhanced selectors for date elements
            date_selectors = [
                "//span[@class='time' and @data-l10n-id='start_date']",
                "//input[@type='date']",
                "//div[contains(@class, 'date-picker')]",
                "//button[contains(@class, 'date-selector')]",
                "//div[contains(@class, 'calendar')]//button",
                "//label[contains(text(), 'Date')]/following-sibling::input",
                "//span[contains(text(), 'Date Range')]",
                "//div[contains(@class, 'date-range-picker')]",
                # More specific selectors for WeatherLink
                "//div[contains(@class, 'date-selector')]",
                "//button[contains(@class, 'btn-date')]",
                "//div[contains(@class, 'date-filter')]",
                "//input[contains(@id, 'date')]",
                "//input[contains(@name, 'date')]",
                "//div[contains(@class, 'datepicker')]",
                # Try to find by aria attributes
                "//*[@aria-label='Date picker' or @aria-label='Calendar' or @aria-label='Date']",
                # Try to find by placeholder
                "//input[@placeholder='Date' or contains(@placeholder, 'date') or contains(@placeholder, 'calendar')]"
            ]
            
            date_element = None
            used_selector = None
            
            # Try each selector to find date control
            for selector in date_selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, selector)
                    if elements:
                        # Try to filter for visible elements only
                        visible_elements = [e for e in elements if e.is_displayed()]
                        if visible_elements:
                            date_element = visible_elements[0]
                        else:
                            date_element = elements[0]
                        used_selector = selector
                        print(f"Found date element with selector: {selector}")
                        break
                except Exception as selector_error:
                    print(f"Error with selector {selector}: {selector_error}")
                    continue
            
            if not date_element:
                print("Could not find any date selector element on the page")
                self.take_screenshot("date_selector_not_found")
                
                # Try looking for any element that might be a date
                try:
                    # Use page source to find potential date elements
                    page_source = self.driver.page_source
                    if "date" in page_source.lower() or "calendar" in page_source.lower():
                        print("Page contains date-related text, trying JavaScript approach...")
                        
                        # Try JavaScript to find potential date inputs
                        date_elements = self.driver.execute_script("""
                            return Array.from(document.querySelectorAll('input,button,div,span'))
                                .filter(el => {
                                    const text = (el.textContent || '').toLowerCase();
                                    const id = (el.id || '').toLowerCase();
                                    const className = (el.className || '').toLowerCase();
                                    const type = (el.type || '').toLowerCase();
                                    const placeholder = (el.placeholder || '').toLowerCase();
                                    
                                    return (text.includes('date') || id.includes('date') || 
                                            className.includes('date') || type === 'date' ||
                                            placeholder.includes('date') || text.includes('calendar') ||
                                            id.includes('calendar') || className.includes('calendar'));
                                });
                        """)
                        
                        if date_elements and len(date_elements) > 0:
                            date_element = date_elements[0]
                            print(f"Found potential date element using JavaScript: {date_element.tag_name}")
                        else:
                            print("No date elements found using JavaScript approach")
                    else:
                        print("No date-related text found in page source")
                except Exception as js_error:
                    print(f"JavaScript search error: {js_error}")
                
                # If still no date element found, return False
                if not date_element:
                    print("Failed to find any date element - cannot set date")
                    return False
            
            # Try to interact with the date element
            try:
                print(f"Clicking on date selector: {used_selector}")
                self.take_screenshot("before_click_date")
                
                # Try multiple click methods
                click_success = False
                
                # Method 1: Standard click
                try:
                    print("Trying standard click...")
                    date_element.click()
                    time.sleep(1)
                    click_success = True
                except Exception as click_error:
                    print(f"Standard click failed: {click_error}")
                
                # Method 2: JavaScript click if standard click failed
                if not click_success:
                    try:
                        print("Trying JavaScript click...")
                        self.driver.execute_script("arguments[0].click();", date_element)
                        time.sleep(1)
                        click_success = True
                    except Exception as js_click_error:
                        print(f"JavaScript click failed: {js_click_error}")
                
                # Method 3: Try sending Enter key
                if not click_success:
                    try:
                        print("Trying to send Enter key...")
                        from selenium.webdriver.common.keys import Keys
                        date_element.send_keys(Keys.ENTER)
                        time.sleep(1)
                        click_success = True
                    except Exception as keys_error:
                        print(f"Sending keys failed: {keys_error}")
                
                # Method 4: Try Actions class
                if not click_success:
                    try:
                        print("Trying Actions click...")
                        from selenium.webdriver.common.action_chains import ActionChains
                        actions = ActionChains(self.driver)
                        actions.move_to_element(date_element).click().perform()
                        time.sleep(1)
                        click_success = True
                    except Exception as actions_error:
                        print(f"Actions click failed: {actions_error}")
                
                # Method 5: If it's an input field with type="date", try to set directly
                if not click_success and date_element.tag_name.lower() == "input":
                    try:
                        input_type = date_element.get_attribute("type")
                        if input_type and input_type.lower() == "date":
                            print("Found date input field, setting value directly...")
                            # Format the date as YYYY-MM-DD for HTML date inputs
                            formatted_date = f"{current_year}-{current_month:02d}-{current_day:02d}"
                            # First clear the field
                            date_element.clear()
                            # Then set the value using JavaScript (more reliable than send_keys)
                            self.driver.execute_script(
                                f"arguments[0].value = '{formatted_date}';", date_element)
                            # Also dispatch change event to ensure the page updates
                            self.driver.execute_script(
                                "arguments[0].dispatchEvent(new Event('change', { 'bubbles': true }));", 
                                date_element)
                            time.sleep(1)
                            click_success = True
                            
                            # Since we directly set the value, we can skip date picker interaction
                            print(f"Date value set directly to {formatted_date}")
                            self.take_screenshot("date_set_directly")
                            return True
                    except Exception as input_error:
                        print(f"Setting input value failed: {input_error}")
                
                if not click_success:
                    print("All click methods failed, cannot proceed with date selection")
                    return False
                
                self.take_screenshot("after_click_date")
                
                # Now look for date picker elements that might have appeared
                date_picker_elements = [
                    "//div[contains(@class, 'calendar')]",
                    "//div[contains(@class, 'datepicker')]",
                    "//div[contains(@class, 'date-picker')]",
                    "//div[contains(@class, 'picker-dropdown')]",
                    "//div[contains(@class, 'date-selector-dropdown')]",
                    "//table[contains(@class, 'calendar')]",
                    "//div[@role='dialog' and contains(@class, 'date')]"
                ]
                
                date_picker = None
                for picker_selector in date_picker_elements:
                    try:
                        elements = self.driver.find_elements(By.XPATH, picker_selector)
                        visible_elements = [e for e in elements if e.is_displayed()]
                        if visible_elements:
                            date_picker = visible_elements[0]
                            print(f"Found date picker with selector: {picker_selector}")
                            break
                    except Exception as picker_error:
                        print(f"Error finding picker with selector {picker_selector}: {picker_error}")
                        continue
                
                if date_picker:
                    print("Date picker is open, trying to set current date")
                    
                    # IMPORTANT: NEW APPROACH - Set year, then month, then day in that specific order
                    print("\n*** SETTING DATE IN SPECIFIC ORDER: YEAR → MONTH → DAY ***\n")
                    
                    # -------- YEAR SELECTION --------
                    year_set = False
                    try:
                        # First click on the datepicker switch to access the month/year view
                        datepicker_switch = self.driver.find_element(By.XPATH, 
                            "//div[contains(@class, 'datepicker')]//th[contains(@class, 'datepicker-switch')]")
                        
                        print(f"Found datepicker switch: {datepicker_switch.text}")
                        
                        # Check if we're already in year or month view
                        if "April 2025" in datepicker_switch.text or any(month_name in datepicker_switch.text for month_name in month_names):
                            # We're in day view with month and year shown - need to click to get to month view
                            print("Currently in day view, clicking to access month view...")
                            datepicker_switch.click()
                            time.sleep(1)
                            
                            # Now click again to go from month view to year view
                            datepicker_switch = self.driver.find_element(By.XPATH, 
                                "//div[contains(@class, 'datepicker')]//th[contains(@class, 'datepicker-switch')]")
                            print(f"Now in month view, switch shows: {datepicker_switch.text}")
                            if str(current_year) in datepicker_switch.text:
                                # We're in month view for the correct year
                                print(f"Month view is showing correct year: {current_year}")
                                year_set = True
                            else:
                                # Click to get to year view
                                print("Clicking again to access year view...")
                                datepicker_switch.click()
                                time.sleep(1)
                        elif str(current_year) in datepicker_switch.text:
                            # We're already in the year view or month view for the correct year
                            print(f"Already showing correct year: {current_year}")
                            year_set = True
                        else:
                            # First click might be needed to get out of day view
                            datepicker_switch.click()
                            time.sleep(1)
                            
                            # Get the switch text again
                            datepicker_switch = self.driver.find_element(By.XPATH, 
                                "//div[contains(@class, 'datepicker')]//th[contains(@class, 'datepicker-switch')]")
                            if not str(current_year) in datepicker_switch.text:
                                # Need to click again to get to year view
                                datepicker_switch.click()
                                time.sleep(1)
                        
                        # If year isn't set yet, look for year span and click it
                        if not year_set:
                            # Look for the year span inside years view
                            year_spans = self.driver.find_elements(By.XPATH, 
                                f"//div[contains(@class, 'datepicker-years')]//span[contains(text(), '{current_year}')]")
                            
                            if year_spans:
                                print(f"Found year span for {current_year}, clicking...")
                                year_spans[0].click()
                                time.sleep(1)
                                year_set = True
                            else:
                                print(f"Year {current_year} not found in year view.")
                                
                                # Navigate through years if needed using next/prev buttons
                                year_view_header = self.driver.find_element(By.XPATH, 
                                    "//div[contains(@class, 'datepicker-years')]//th[contains(@class, 'datepicker-switch')]")
                                year_range_text = year_view_header.text  # e.g., "2020-2029"
                                
                                print(f"Year view range: {year_range_text}")
                                
                                if year_range_text:
                                    try:
                                        # Parse the year range
                                        years = [int(y) for y in year_range_text.split("-")]
                                        range_start = years[0]
                                        range_end = years[1]
                                        
                                        # Determine if we need to navigate
                                        if current_year < range_start:
                                            # Need to go back
                                            print(f"Target year {current_year} is before current range {range_start}-{range_end}, clicking prev...")
                                            prev_button = self.driver.find_element(By.XPATH, 
                                                "//div[contains(@class, 'datepicker-years')]//th[contains(@class, 'prev')]")
                                            prev_button.click()
                                            time.sleep(1)
                                        elif current_year > range_end:
                                            # Need to go forward
                                            print(f"Target year {current_year} is after current range {range_start}-{range_end}, clicking next...")
                                            next_button = self.driver.find_element(By.XPATH, 
                                                "//div[contains(@class, 'datepicker-years')]//th[contains(@class, 'next')]")
                                            next_button.click()
                                            time.sleep(1)
                                        
                                        # Now look for the year again
                                        year_spans = self.driver.find_elements(By.XPATH, 
                                            f"//div[contains(@class, 'datepicker-years')]//span[contains(text(), '{current_year}')]")
                                        
                                        if year_spans:
                                            print(f"Found year {current_year} after navigation, clicking...")
                                            year_spans[0].click()
                                            time.sleep(1)
                                            year_set = True
                                    except Exception as parse_error:
                                        print(f"Error parsing year range: {parse_error}")
                    except Exception as year_error:
                        print(f"Error setting year: {year_error}")
                    
                    # -------- MONTH SELECTION --------
                    month_set = False
                    try:
                        # Based on the HTML, we're looking for a month element with specific class
                        # <span class="month focused active">Mar</span> vs <span class="month">Apr</span>
                        
                        # First, check if we're in month view
                        month_container = self.driver.find_elements(By.XPATH, 
                            "//div[contains(@class, 'datepicker-months')]//td[contains(@colspan, '7')]")
                        
                        if month_container:
                            print("Found month selection view")
                            
                            # Look for the specific month using both the full name and abbreviated name
                            month_abbreviated = current_month_name[:3]  # First 3 letters (e.g., "Apr")
                            
                            month_span = self.driver.find_elements(By.XPATH, 
                                f"//div[contains(@class, 'datepicker-months')]//span[@class='month' and (text()='{current_month_name}' or text()='{month_abbreviated}')]")
                            
                            if month_span:
                                print(f"Found month span for {current_month_name}, clicking...")
                                month_span[0].click()
                                time.sleep(1)
                                month_set = True
                            else:
                                print(f"Could not find span for month {current_month_name} or {month_abbreviated}")
                                
                                # Fallback: try to find by position (months are 0-indexed)
                                try:
                                    all_months = self.driver.find_elements(By.XPATH, 
                                        "//div[contains(@class, 'datepicker-months')]//span[contains(@class, 'month')]")
                                    
                                    if len(all_months) >= current_month:
                                        print(f"Selecting month {current_month_name} by position ({current_month-1})...")
                                        all_months[current_month-1].click()
                                        time.sleep(1)
                                        month_set = True
                                except Exception as month_pos_error:
                                    print(f"Error selecting month by position: {month_pos_error}")
                        else:
                            print("Month selection view not found")
                            
                            # Check if we're in day view with the correct month already
                            day_view_header = self.driver.find_elements(By.XPATH, 
                                "//div[contains(@class, 'datepicker-days')]//th[contains(@class, 'datepicker-switch')]")
                            
                            if day_view_header and current_month_name in day_view_header[0].text:
                                print(f"Already in correct month: {current_month_name} {current_year}")
                                month_set = True
                            elif day_view_header:
                                # We're in day view but wrong month, click to get to month view
                                print(f"In day view but wrong month: {day_view_header[0].text}, clicking to access month view...")
                                day_view_header[0].click()
                                time.sleep(1)
                                
                                # Now try to find and click the month again
                                month_span = self.driver.find_elements(By.XPATH, 
                                    f"//div[contains(@class, 'datepicker-months')]//span[@class='month' and (text()='{current_month_name}' or text()='{month_abbreviated}')]")
                                
                                if month_span:
                                    print(f"Found month span for {current_month_name}, clicking...")
                                    month_span[0].click()
                                    time.sleep(1)
                                    month_set = True
                    except Exception as month_error:
                        print(f"Error setting month: {month_error}")
                    
                    # -------- DAY SELECTION --------
                    day_set = False
                    try:
                        # Based on the HTML, we'll target day elements in the current month
                        # Avoid days with class "old" or "new" which are from adjacent months
                        
                        # Look for the day but exclude old or new days
                        day_element = self.driver.find_elements(By.XPATH, 
                            f"//div[contains(@class, 'datepicker-days')]//td[text()='{current_day}' and not(contains(@class, 'old')) and not(contains(@class, 'new'))]")
                        
                        if day_element:
                            print(f"Found day element for {current_day} in current month, clicking...")
                            day_element[0].click()
                            time.sleep(1)
                            day_set = True
                        else:
                            # Try to find any day element with the correct number
                            all_day_elements = self.driver.find_elements(By.XPATH, 
                                f"//div[contains(@class, 'datepicker-days')]//td[text()='{current_day}']")
                            
                            if all_day_elements:
                                print(f"Found day {current_day} (may not be in current month), clicking...")
                                all_day_elements[0].click()
                                time.sleep(1)
                                day_set = True
                    except Exception as day_error:
                        print(f"Error setting day: {day_error}")
                    
                    if not (year_set and month_set and day_set):
                        print(f"WARNING: Could not fully set date to {current_month_name} {current_day}, {current_year}")
                    else:
                        print(f"Successfully set date to {current_month_name} {current_day}, {current_year}")
                        
                    # Close the date picker if it's still open
                    try:
                        # Look for any "apply" button or click outside to close
                        apply_button = self.driver.find_elements(By.XPATH, 
                            "//button[contains(@class, 'apply') or contains(text(), 'Apply')]")
                        if apply_button:
                            apply_button[0].click()
                            print("Clicked apply button to close date picker")
                        else:
                            # Try clicking outside the date picker to close it
                            actions = ActionChains(self.driver)
                            actions.move_by_offset(10, 10).click().perform()
                            print("Clicked outside date picker to close it")
                    except Exception as close_error:
                        print(f"Error closing date picker: {close_error}")
                        # Continue anyway as this is not critical
                
                # Try to verify that the date is now set to current
                print("Checking if the date was updated correctly")
                self.take_screenshot("after_date_selection")
                
                # Look for date display elements to verify current date
                date_display_selectors = [
                    "//span[@class='time']",
                    "//input[@type='date']",
                    "//div[contains(@class, 'selected-date')]",
                    "//button[contains(@class, 'date-display')]",
                    "//span[contains(@class, 'date-display')]",
                    "//input[contains(@class, 'date') or @type='date']",
                    # Add the original date element we interacted with
                    used_selector if used_selector else None
                ]
                
                date_verified = False
                for selector in date_display_selectors:
                    if not selector:
                        continue
                        
                    try:
                        elements = self.driver.find_elements(By.XPATH, selector)
                        visible_elements = [e for e in elements if e.is_displayed()]
                        
                        if visible_elements:
                            element = visible_elements[0]
                            # Get text content or value
                            date_text = element.text
                            if not date_text and element.tag_name == "input":
                                date_text = element.get_attribute("value")
                                
                            print(f"Found date display: {date_text}")
                            
                            if not date_text:
                                print("Date text is empty, trying next selector")
                                continue
                                
                            # Prepare current date components for comparison
                            current_day_str = str(current_day)
                            current_month_str = str(current_month)
                            current_year_str = str(current_year)
                            current_year_short = str(current_year)[-2:]  # Last 2 digits of year
                            
                            # If day or month is single digit, also create padded versions (e.g., 01 instead of 1)
                            padded_day = f"0{current_day}" if current_day < 10 else current_day_str
                            padded_month = f"0{current_month}" if current_month < 10 else current_month_str
                            
                            # Extract all date-like patterns from the text
                            import re
                            # Look for date patterns like MM/DD/YY, DD/MM/YY, etc.
                            date_patterns = re.findall(r'(\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4})', date_text)
                            
                            if date_patterns:
                                print(f"Found date patterns in text: {date_patterns}")
                                for pattern in date_patterns:
                                    # Try to parse this pattern
                                    date_parts = re.split(r'[-/\.]', pattern)
                                    if len(date_parts) == 3:
                                        # The key change here - look specifically for year-month-day format
                                        # since the user wants year first, then month, then day
                                        
                                        # Check YYYY-MM-DD format (year first)
                                        if (date_parts[0] == current_year_str or date_parts[0] == current_year_short) and \
                                           (date_parts[1] == current_month_str or date_parts[1] == padded_month) and \
                                           (date_parts[2] == current_day_str or date_parts[2] == padded_day):
                                            print(f"Verified correct date in YYYY-MM-DD format: {pattern}")
                                            date_verified = True
                                            break
                                            
                                        # Also check MM/DD/YY format (most common in US)
                                        elif (date_parts[0] == current_month_str or date_parts[0] == padded_month) and \
                                             (date_parts[1] == current_day_str or date_parts[1] == padded_day) and \
                                             (date_parts[2] == current_year_str or date_parts[2] == current_year_short):
                                            print(f"Verified correct date in MM/DD/YY format: {pattern}")
                                            date_verified = True
                                            break
                                            
                                        # If found DD/MM/YY pattern, this might be problematic
                                        elif (date_parts[0] == current_day_str or date_parts[0] == padded_day) and \
                                             (date_parts[1] == current_month_str or date_parts[1] == padded_month) and \
                                             (date_parts[2] == current_year_str or date_parts[2] == current_year_short):
                                            print(f"WARNING: Found date in DD/MM/YY format ({pattern}), but we wanted YYYY-MM-DD")
                                            print(f"This may indicate a problem with the date selection")
                                            # Don't set as verified - we need to continue trying
                                                
                                        # Print details for debugging
                                        print(f"Date pattern {pattern} analysis:")
                                        print(f"Parts: {date_parts}")
                                        print(f"Expected: Year={current_year_str}/{current_year_short}, Month={current_month_str}/{padded_month}, Day={current_day_str}/{padded_day}")
                            
                            # If no date pattern matched, check for presence of date components and month name
                            if not date_verified:
                                # Check if month name is present along with day and year
                                if (current_month_name in date_text or current_month_name.lower() in date_text.lower()) and \
                                   (current_day_str in date_text or padded_day in date_text) and \
                                   (current_year_str in date_text or current_year_short in date_text):
                                    print(f"Verified date by components: {current_month_name}, {current_day}, {current_year}")
                                    date_verified = True
                                else:
                                    # If we get here, the date text doesn't contain the expected date in any recognizable format
                                    print(f"Date text '{date_text}' doesn't match expected date: {current_month}/{current_day}/{current_year} ({current_month_name})")
                            
                            if date_verified:
                                break
                    except Exception as verify_error:
                        print(f"Error verifying date with selector {selector}: {verify_error}")
                        continue
                
                if not date_verified:
                    print("\nWARNING: Could not verify date was set correctly to current system date")
                    print(f"Expected: {current_month}/{current_day}/{current_year} ({current_month_name}) in YYYY-MM-DD format")
                    print("The displayed date may not reflect the current system date or may be in an unexpected format")
                    # Take one final screenshot showing the current state
                    self.take_screenshot("date_verification_failed")
                    # Since we made our best effort, we'll return False to indicate this operation wasn't successful
                    return False
                else:
                    print(f"\nSUCCESS: Date verified as correctly set to current system date: {current_month}/{current_day}/{current_year}")
                
                return True
                
            except Exception as e:
                print(f"Error setting date: {e}")
                self.take_screenshot("set_date_error")
                # Return False to indicate failure, but the calling function can decide to continue
                return False
        
        except Exception as e:
            print(f"Error setting date: {e}")
            self.take_screenshot("set_date_error")
            # Return False to indicate failure, but the calling function can decide to continue
            return False
    
    def take_screenshot(self, name):
        """Take a screenshot and save it with the given name."""
        if self.debug:
            try:
                # Create screenshots directory if it doesn't exist
                if not os.path.exists("screenshots"):
                    os.makedirs("screenshots")
                
                # Generate filename with timestamp
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"screenshots/{name}_{timestamp}.png"
                
                # Take screenshot
                self.driver.save_screenshot(filename)
                print(f"Screenshot saved: {filename}")
            except Exception as e:
                print(f"Error taking screenshot: {e}")
    
    def navigate_to_data_page(self):
        """Navigate to the data page."""
        try:
            print("Navigating to data page...")
            self.take_screenshot("before_navigation")
            
            max_retries = 3
            retry_delay = 5  # seconds
            
            for attempt in range(max_retries):
                try:
                    # Check if we need to log in again
                    if "login" in self.driver.current_url.lower():
                        print("Detected login page, attempting to log in again...")
                        if not self.login():
                            print("Failed to log in")
                            return False
                        print("Login successful")
                        time.sleep(3)  # Wait for login to complete
                    
                    # Try direct URL navigation with the specific URL
                    print("Trying direct URL navigation to the data page...")
                    specific_url = f"{self.url}browse/2bd0cbc6-a874-441f-99f4-2410a8143886"
                    self.driver.get(specific_url)
                    print(f"Navigated directly to: {specific_url}")
                    time.sleep(5)  # Give more time for the page to load
                    
                    # Take a screenshot of the data page
                    self.take_screenshot("data_page")
                    
                    # Log the current URL so we can see where we ended up
                    print(f"Current URL after navigation: {self.driver.current_url}")
                    
                    # Check if we ended up on the login page again
                    if "login" in self.driver.current_url.lower():
                        print("Redirected to login page, attempting to log in again...")
                        if not self.login():
                            print("Failed to log in")
                            return False
                        print("Login successful")
                        time.sleep(3)  # Wait for login to complete
                        
                        # Try direct URL navigation once more
                        self.driver.get(specific_url)
                        print(f"Navigated directly to: {specific_url}")
                        time.sleep(5)
                        
                        # Check if we're still on the login page
                        if "login" in self.driver.current_url.lower():
                            print("Still on login page, navigation failed")
                            if attempt < max_retries - 1:
                                print(f"Will retry in {retry_delay} seconds...")
                                time.sleep(retry_delay)
                                continue
                            else:
                                print("Max retries reached. Navigation failed.")
                                return False
                    
                    # Check for successful navigation by looking for typical elements on the data page
                    try:
                        # Look for elements that would be on the data page
                        data_indicators = [
                            "//table[contains(@class, 'data-table')]",
                            "//div[contains(@class, 'data-view')]",
                            "//h1[contains(text(), 'Data')]",
                            "//a[@id='export-data']"
                        ]
                        
                        navigation_success = False
                        for indicator in data_indicators:
                            try:
                                element = WebDriverWait(self.driver, 5).until(
                                    EC.presence_of_element_located((By.XPATH, indicator))
                                )
                                print(f"Found data page indicator: {indicator}")
                                navigation_success = True
                                break
                            except:
                                continue
                        
                        if navigation_success:
                            print("Successfully navigated to data page")
                            return True
                        else:
                            print("Could not confirm successful navigation to data page")
                            # Take another screenshot to debug
                            self.take_screenshot("data_page_confirmation_failed")
                            # Continue anyway - the next steps will fail if this isn't right
                            return True
                    except Exception as confirm_error:
                        print(f"Error confirming navigation: {confirm_error}")
                        # Continue anyway
                        return True
                    
                except Exception as e:
                    print(f"Error during navigation attempt {attempt + 1}: {e}")
                    if attempt < max_retries - 1:
                        print(f"Waiting {retry_delay} seconds before retrying...")
                        time.sleep(retry_delay)
                    else:
                        print("Max retries reached. Navigation failed.")
                        self.take_screenshot("navigation_error")
                        return False
            
        except Exception as e:
            print(f"Error navigating to data page: {e}")
            self.take_screenshot("navigation_error")
            return False
    
    def extract_weather_data(self):
        """Extract weather data from the current page."""
        try:
            print("Extracting weather data...")
            self.take_screenshot("before_extraction")
            
            # Wait for the data table to be present
            table = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table.data-table"))
            )
            
            # Get all rows from the table
            rows = table.find_elements(By.TAG_NAME, "tr")
            if not rows:
                print("No rows found in the table")
                return None
                
            # Get headers from the first row
            headers = [cell.text.strip() for cell in rows[0].find_elements(By.TAG_NAME, "th")]
            if not headers:
                print("No headers found in the table")
                return None
                
            print(f"Found {len(headers)} columns: {headers}")
            
            # Extract data from remaining rows
            data = []
            for row in rows[1:]:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) != len(headers):
                    print(f"Warning: Row has {len(cells)} cells but expected {len(headers)}")
                    continue
                    
                row_data = []
                for cell in cells:
                    row_data.append(cell.text.strip())
                if row_data:  # Only add non-empty rows
                    data.append(row_data)
                    
            print(f"Extracted {len(data)} rows of data")
            
            # Create DataFrame from extracted data
            if data:
                df = pd.DataFrame(data, columns=headers)
                print(f"Created DataFrame with shape: {df.shape}")
                return df
            else:
                print("No data rows found")
                return None
            
        except Exception as e:
            print(f"Error extracting data: {e}")
            self.take_screenshot("extraction_error")
            return None
    
    def export_data(self, email=None):
        """Export data by clicking the export button and sending to email."""
        try:
            # Use provided email or default to the one set in constructor
            if email is None:
                email = self.export_email
                
            print(f"Exporting data to email: {email}...")
            self.take_screenshot("before_export")
            
            # Check if we're on the data page, if not, try to navigate there
            if not ("browse" in self.driver.current_url.lower() or "data" in self.driver.current_url.lower()):
                print("Not on data page, attempting to navigate to data page...")
                if not self.navigate_to_data_page():
                    print("Failed to navigate to data page")
                    return False
                
            # Check if we need to log in again
            if "login" in self.driver.current_url.lower():
                print("Redirected to login page, attempting to log in again...")
                if not self.login():
                    print("Failed to log in")
                    return False
                print("Login successful")
                time.sleep(3)  # Wait for login to complete
                
                # Navigate to data page again
                if not self.navigate_to_data_page():
                    print("Failed to navigate to data page after login")
                    return False
            
            # Check if there's a modal dialog blocking the UI and close it
            try:
                print("Checking for modal dialogs that might block the UI...")
                modal = self.driver.find_element(By.ID, "modal-config")
                if modal.is_displayed():
                    print("Found a modal dialog blocking the UI. Attempting to close it...")
                    # Try to find close button or X button
                    close_buttons = self.driver.find_elements(By.XPATH, 
                        "//div[@id='modal-config']//button[contains(@class, 'close') or contains(@class, 'btn-close')]")
                    
                    if close_buttons:
                        print("Found close button for the modal")
                        close_buttons[0].click()
                        print("Clicked close button")
                    else:
                        # Try to click anywhere outside the modal to close it
                        print("No close button found. Trying to click outside the modal...")
                        actions = ActionChains(self.driver)
                        actions.move_by_offset(10, 10).click().perform()
                        
                    # Wait for modal to disappear
                    print("Waiting for modal to disappear...")
                    WebDriverWait(self.driver, 5).until_not(
                        EC.visibility_of_element_located((By.ID, "modal-config"))
                    )
                    print("Modal is no longer visible")
                else:
                    print("Modal is present but not displayed")
            except Exception as e:
                print(f"No blocking modal found or error handling modal: {e}")
            
            # Wait for the export button to be clickable
            print("Looking for export button...")
            time.sleep(2)  # Give the page a moment to settle after modal handling
            
            # Try different selectors for the export button
            export_button = None
            export_selectors = [
                "a#export-data",
                "a[id='export-data']",
                "a.export-button",
                "button.export-button",
                "a[title='Export Data']",
                "a[href='#export']",
                "//a[contains(text(), 'Export')]",
                "//button[contains(text(), 'Export')]"
            ]
            
            for selector in export_selectors:
                try:
                    # Determine if it's a CSS or XPath selector
                    if selector.startswith("//"):
                        export_button = WebDriverWait(self.driver, 3).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                    else:
                        export_button = WebDriverWait(self.driver, 3).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                        )
                    if export_button:
                        print(f"Found export button with selector: {selector}")
                        break
                except:
                    continue
            
            if not export_button:
                print("Could not find any export button")
                self.take_screenshot("export_button_not_found")
                return False
            
            # Try alternative methods to click the button if direct click fails
            try:
                # First try: standard click
                print("Clicking export button...")
                export_button.click()
            except Exception as click_error:
                print(f"Standard click failed: {click_error}")
                print("Trying JavaScript click...")
                self.driver.execute_script("arguments[0].click();", export_button)
                
            time.sleep(2)
            self.take_screenshot("export_form_opened")
            
            # Wait for the email input field
            print("Looking for email input field...")
            email_field = None
            try:
                email_field = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input#email.form-control"))
                )
            except:
                # Try alternate selectors
                try:
                    email_field = WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@type='email' or @id='email' or @name='email']"))
                    )
                except:
                    print("Could not find email input field")
                    self.take_screenshot("email_field_not_found")
                    return False
                    
            print("Found email input field")
            
            # Clear any existing text and enter the email
            email_field.clear()
            print(f"Entering email: {email}")
            email_field.send_keys(email)
            time.sleep(1)
            self.take_screenshot("email_entered")
            
            # Look for the Send button or Export button
            print("Looking for Send/Export button...")
            send_button = None
            send_selectors = [
                "button#js-updateBtn.export-button",
                "button.export-button",
                "input[type='submit']",
                "button[type='submit']",
                "//button[contains(text(), 'Send')]",
                "//button[contains(text(), 'Export')]",
                "//input[@value='Export']"
            ]
            
            for selector in send_selectors:
                try:
                    # Determine if it's a CSS or XPath selector
                    if selector.startswith("//"):
                        send_button = WebDriverWait(self.driver, 3).until(
                            EC.element_to_be_clickable((By.XPATH, selector))
                        )
                    else:
                        send_button = WebDriverWait(self.driver, 3).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                        )
                    if send_button:
                        print(f"Found send/export button with selector: {selector}")
                        break
                except:
                    continue
                    
            if not send_button:
                print("Could not find any send/export button")
                self.take_screenshot("send_button_not_found")
                return False
            
            print("Clicking Send button...")
            try:
                # First try: standard click
                send_button.click()
            except Exception as click_error:
                print(f"Standard click failed: {click_error}")
                print("Trying JavaScript click...")
                self.driver.execute_script("arguments[0].click();", send_button)
                
            time.sleep(3)
            self.take_screenshot("after_export_sent")
            
            # Look for confirmation message
            try:
                confirmation = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'alert-success') or contains(text(), 'success') or contains(text(), 'sent')]"))
                )
                print("Found success confirmation message")
            except:
                # If no confirmation found, check for error messages
                try:
                    error_msg = WebDriverWait(self.driver, 2).until(
                        EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'alert-danger') or contains(@class, 'alert-error')]"))
                    )
                    print(f"Error message found: {error_msg.text}")
                    self.take_screenshot("export_error_message")
                except:
                    print("No confirmation or error message found")
                    print("Assuming export request was sent anyway")
            
            print("Data export request sent successfully")
            return True
            
        except Exception as e:
            print(f"Error during data export: {e}")
            self.take_screenshot("export_error")
            return False

    def check_gmail_api(self, max_wait_time=300):
        """
        Check Gmail for the exported data email using the API.
        
        Args:
            max_wait_time: Maximum time to wait for the email (in seconds)
            
        Returns:
            Path to the downloaded CSV file or None if not found
        """
        try:
            print("\nAutomatically checking Gmail using API for the exported data...")
            
            if GmailAPI is None:
                print("GmailAPI module is not available. Please check the import error above.")
                return None
            
            # Initialize the GmailAPI
            try:
                print("Initializing GmailAPI...")
                gmail_api = GmailAPI()
                if not gmail_api:
                    print("GmailAPI initialization returned None")
                    return None
                if not hasattr(gmail_api, 'service'):
                    print("GmailAPI instance has no 'service' attribute")
                    return None
                print("GmailAPI initialized successfully")
            except Exception as init_error:
                print(f"Error initializing GmailAPI: {init_error}")
                print("Please check your credentials.json file and Gmail API setup")
                return None
            
            # Set up for waiting - check multiple times with delays
            start_time = time.time()
            retry_count = 0
            max_retries = 10  # Maximum number of retries
            retry_delay = 30  # Initial delay between retries in seconds
            
            # Store message IDs we've already checked to avoid re-checking the same messages
            checked_message_ids = set()
            
            while time.time() - start_time < max_wait_time and retry_count < max_retries:
                try:
                    print(f"\nChecking for WeatherLink emails (Attempt {retry_count + 1}/{max_retries})...")
                    
                    # Get the messages
                    messages = gmail_api.service.users().messages().list(
                        userId='me',
                        q='from:weatherlink.com',
                        maxResults=10
                    ).execute()
                    
                    if 'messages' in messages and messages['messages']:
                        print(f"Found {len(messages['messages'])} messages from WeatherLink")
                        
                        # Process messages from newest to oldest
                        for msg_info in messages['messages']:
                            msg_id = msg_info['id']
                            
                            # Skip if we've already checked this message
                            if msg_id in checked_message_ids:
                                continue
                            
                            checked_message_ids.add(msg_id)
                            
                            try:
                                print(f"Examining message {msg_id}...")
                                msg = gmail_api.service.users().messages().get(
                                    userId='me',
                                    id=msg_id,
                                    format='full'
                                ).execute()
                                
                                # Get timestamp and check if this is a recent message
                                headers = msg['payload']['headers']
                                date_header = next((h for h in headers if h['name'].lower() == 'date'), None)
                                
                                if date_header:
                                    print(f"Message date: {date_header['value']}")
                                
                                # Extract the S3 URL from the email content
                                s3_url = None
                                
                                # Check message structure - look in parts first
                                if 'parts' in msg['payload']:
                                    print("Message has parts structure - analyzing...")
                                    for part_idx, part in enumerate(msg['payload']['parts']):
                                        print(f"Examining part {part_idx+1} of {len(msg['payload']['parts'])}: {part.get('mimeType', 'unknown type')}")
                                        if part['mimeType'] == 'text/html':
                                            data = part['body'].get('data', '')
                                            if data:
                                                import base64
                                                html = base64.urlsafe_b64decode(data).decode('utf-8')
                                                print(f"HTML content preview (first 200 chars): {html[:200]}...")
                                                
                                                # Look for the download link in the HTML - try multiple patterns
                                                import re
                                                # First attempt with specific pattern
                                                s3_url_pattern = r'https://s3\.amazonaws\.com/export-wl2-live\.weatherlink\.com/data/[^"]+\.csv'
                                                match = re.search(s3_url_pattern, html)
                                                if match:
                                                    s3_url = match.group(0)
                                                    print(f"Found S3 URL with specific pattern: {s3_url}")
                                                    break
                                                
                                                # Second attempt with more general pattern
                                                s3_url_pattern_alt = r'https://s3\.amazonaws\.com/[^"]+\.csv'
                                                match = re.search(s3_url_pattern_alt, html)
                                                if match:
                                                    s3_url = match.group(0)
                                                    print(f"Found S3 URL with general pattern: {s3_url}")
                                                    break
                                                
                                                # Third attempt looking specifically for href attributes
                                                href_pattern = r'href="(https://s3\.amazonaws\.com/[^"]+\.csv)"'
                                                match = re.search(href_pattern, html)
                                                if match:
                                                    s3_url = match.group(1)
                                                    print(f"Found S3 URL in href attribute: {s3_url}")
                                                    break
                                                    
                                                # If still not found, search for any link with .csv extension
                                                csv_link_pattern = r'href="([^"]+\.csv)"'
                                                match = re.search(csv_link_pattern, html)
                                                if match:
                                                    s3_url = match.group(1)
                                                    print(f"Found generic CSV link: {s3_url}")
                                                    break
                                                    
                                                print("No S3 URL patterns matched in HTML content")
                                            else:
                                                print("Part has no data content")
                                # Check for body directly in the message if no parts or S3 URL not found in parts
                                elif 'body' in msg['payload'] and not s3_url:
                                    print("Checking message body directly...")
                                    data = msg['payload']['body'].get('data', '')
                                    if data:
                                        import base64
                                        html = base64.urlsafe_b64decode(data).decode('utf-8')
                                        print(f"Body content preview (first 200 chars): {html[:200]}...")
                                        
                                        # Same patterns as above
                                        import re
                                        # Try all patterns
                                        patterns = [
                                            r'https://s3\.amazonaws\.com/export-wl2-live\.weatherlink\.com/data/[^"]+\.csv',
                                            r'https://s3\.amazonaws\.com/[^"]+\.csv',
                                            r'href="(https://s3\.amazonaws\.com/[^"]+\.csv)"',
                                            r'href="([^"]+\.csv)"'
                                        ]
                                        
                                        for pattern in patterns:
                                            match = re.search(pattern, html)
                                            if match:
                                                # If it's a capturing group, get group 1, otherwise get the whole match
                                                s3_url = match.group(1) if '(' in pattern else match.group(0)
                                                print(f"Found S3 URL in body using pattern {pattern}: {s3_url}")
                                                break
                                
                                if s3_url:
                                    print(f"\nFound S3 URL: {s3_url}")
                                    
                                    # Create a directory for saved files if it doesn't exist
                                    if not os.path.exists('saved_emails'):
                                        os.makedirs('saved_emails')
                                    
                                    # Extract filename from URL
                                    filename = s3_url.split('/')[-1]
                                    filepath = f"saved_emails/{filename}"
                                    
                                    # Download the file directly using requests
                                    print(f"Downloading file from {s3_url}...")
                                    try:
                                        import requests
                                        response = requests.get(s3_url)
                                        if response.status_code == 200:
                                            with open(filepath, 'wb') as f:
                                                f.write(response.content)
                                            print(f"Successfully downloaded file: {filepath}")
                                            
                                            # Try to clean the CSV data
                                            try:
                                                print("\nCleaning the data...")
                                                import pandas as pd
                                                
                                                # Try different approaches for reading the file
                                                try:
                                                    # Approach 1: Skip the first 5 rows directly during loading
                                                    print("Attempting to read CSV, skipping first 5 rows...")
                                                    df = pd.read_csv(filepath, skiprows=5, encoding='utf-8', 
                                                                    on_bad_lines='skip', low_memory=False)
                                                except UnicodeDecodeError:
                                                    # If UTF-8 fails, try with latin-1 encoding
                                                    print("UTF-8 encoding failed, trying with latin-1 encoding...")
                                                    df = pd.read_csv(filepath, skiprows=5, encoding='latin-1', 
                                                                    on_bad_lines='skip', low_memory=False)
                                                except pd.errors.ParserError as pe:
                                                    print(f"Parser error: {pe}")
                                                    print("Trying with different delimiter...")
                                                    # Try with a different delimiter
                                                    df = pd.read_csv(filepath, skiprows=5, encoding='utf-8', 
                                                                    on_bad_lines='skip', low_memory=False, delimiter=',', 
                                                                    quotechar='"')
                                                except Exception as e:
                                                    # If standard approach fails, try a more manual approach
                                                    print(f"Standard CSV reading failed: {e}")
                                                    print("Trying manual file reading approach...")
                                                    
                                                    # Read raw lines, skip the first 5, then parse
                                                    with open(filepath, 'rb') as f:
                                                        lines = f.readlines()
                                                    
                                                    # Skip first 5 lines, then join the rest
                                                    data_lines = lines[5:]
                                                    
                                                    # Write to a temporary file
                                                    temp_file = 'temp_clean_data.csv'
                                                    with open(temp_file, 'wb') as f:
                                                        f.writelines(data_lines)
                                                    
                                                    # Now try to read the clean temp file
                                                    print("Reading from clean temporary file...")
                                                    df = pd.read_csv(temp_file, encoding='latin-1', 
                                                                    on_bad_lines='skip', low_memory=False)
                                                
                                                # Save the cleaned data
                                                print(f"Successfully read CSV data with shape: {df.shape}")
                                                
                                                # Remove specified columns
                                                df = self.clean_csv_data(df)
                                                
                                                # Save the clean data to dataset.csv
                                                df.to_csv('dataset.csv', index=False)
                                                print("Cleaned data saved as dataset.csv")
                                                
                                                # Run clean.py BEFORE returning
                                                success = self.run_clean_script()
                                                print(f"Clean script run {'successfully' if success else 'with errors'}")
                                                
                                                # Now return the dataframe
                                                return (filepath, df)
                                            except Exception as clean_error:
                                                print(f"Warning: Could not clean the CSV data: {clean_error}")
                                                print("Will still return the original downloaded file.")
                                            
                                            # Return just the filepath if cleaning failed
                                            return filepath
                                        else:
                                            print(f"Failed to download file. Status code: {response.status_code}")
                                    except Exception as download_error:
                                        print(f"Error downloading file: {download_error}")
                            except Exception as msg_error:
                                print(f"Error processing message {msg_id}: {msg_error}")
                    else:
                        print("No WeatherLink emails found in this check.")
                
                except Exception as check_error:
                    print(f"Error during email check attempt {retry_count + 1}: {check_error}")
                
                # Increment retry count
                retry_count += 1
                
                # If we haven't found the email yet and haven't exceeded our wait time, wait before retrying
                if time.time() - start_time < max_wait_time and retry_count < max_retries:
                    wait_time = min(retry_delay, max_wait_time - (time.time() - start_time))
                    print(f"\nNo S3 URL found yet. Waiting {int(wait_time)} seconds before checking again...")
                    print(f"Time elapsed: {int(time.time() - start_time)} seconds out of {max_wait_time} max wait time")
                    time.sleep(wait_time)
            
            print("\nMaximum wait time or retry limit reached. No valid WeatherLink export email found.")
            return None
    
        except Exception as e:
            error_msg = str(e)
            if "access_denied" in error_msg or "verification process" in error_msg:
                print("\n" + "="*80)
                print("GMAIL API ACCESS ERROR")
                print("="*80)
                print("The Gmail API access is currently blocked because the application hasn't been verified.")
                print("\nTo fix this, you need to:")
                print("1. Go to https://console.cloud.google.com")
                print("2. Select your project")
                print("3. Go to 'APIs & Services' > 'OAuth consent screen'")
                print("4. Under 'Test users', add your email address: teamdavcast@gmail.com")
                print("\nAlternatively, you can run the script without the Gmail API:")
                print("python weatherlink.py --debug --check-email --no-api")
                print("="*80 + "\n")
            else:
                print(f"Error checking Gmail using API: {e}")
            
            return None

    def is_logged_in(self):
        """Check if already logged in to the website."""
        try:
            print("Checking login status...")
            self.take_screenshot("login_status_check")
            
            # Look for elements that indicate being logged in
            logged_in_indicators = [
                # Common logout links
                "//a[contains(text(), 'Logout') or contains(@href, 'logout')]",
                "//a[contains(text(), 'Sign Out') or contains(@href, 'signout')]",
                "//a[contains(text(), 'Log Out')]",
                
                # Profile or account links that are typically only shown when logged in
                "//a[contains(text(), 'Profile') or contains(text(), 'Account')]",
                "//a[contains(text(), 'My Account') or contains(text(), 'Dashboard')]",
                
                # Username displays
                "//span[contains(@class, 'username')]",
                "//div[contains(@class, 'user-info')]",
                "//div[contains(@class, 'account-info')]",
                
                # Dashboard elements
                "//h1[contains(text(), 'Dashboard')]",
                "//div[contains(@class, 'dashboard')]"
            ]
            
            # Check each indicator
            for indicator in logged_in_indicators:
                elements = self.driver.find_elements(By.XPATH, indicator)
                if elements:
                    print(f"Found login indicator: {indicator}")
                    return True
                    
            # Check if we're still on the login page
            login_indicators = [
                "//input[@type='password']",
                "//button[contains(text(), 'Log In')]",
                "//h1[contains(text(), 'Login')]",
                "//div[contains(text(), 'Sign In')]"
            ]
            
            for indicator in login_indicators:
                elements = self.driver.find_elements(By.XPATH, indicator)
                if elements:
                    print(f"Still on login page (found {indicator})")
                    return False
                    
            # No clear indication - check the URL
            current_url = self.driver.current_url
            print(f"Current URL: {current_url}")
            
            if "login" in current_url.lower() or "signin" in current_url.lower():
                print("URL indicates we're on a login page")
                return False
                
            # If we've reached a page that's not a login page, we're likely logged in
            print("No login indicators found, but we're not on a login page")
            return True
                
        except Exception as e:
            print(f"Error checking login status: {e}")
            self.take_screenshot("login_status_error")
            return False
        
    def login(self):
        """Login to the WeatherLink website."""
        try:
            print("Attempting to log in...")
            self.take_screenshot("login_page")
            
            # Wait for login form to be visible
            print("Waiting for login form...")
            WebDriverWait(self.driver, 15).until(
                EC.visibility_of_element_located((By.XPATH, "//input[@placeholder='Username' or @name='username' or @id='username']"))
            )
            
            # Locate username input field - try multiple possible selectors
            print("Finding username field...")
            try:
                username_field = self.driver.find_element(By.XPATH, "//input[@placeholder='Username' or @name='username' or @id='username']")
            except:
                # Fallback to regular ID/name attributes
                username_field = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//input[@type='text' and (@id='email' or @name='email' or @name='username')]"))
                )
            
            # Locate password input field
            print("Finding password field...")
            try:
                password_field = self.driver.find_element(By.XPATH, "//input[@type='password' and (@id='password' or @name='password')]")
            except:
                # Fallback to any password field
                password_field = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//input[@type='password']"))
                )
            
            # Get credentials from environment variables
            username = os.getenv("WEATHERLINK_USERNAME")
            password = os.getenv("WEATHERLINK_PASSWORD")
            
            if not username or not password:
                print("Error: Username or password not found in environment variables.")
                print("Please set WEATHERLINK_USERNAME and WEATHERLINK_PASSWORD environment variables.")
                return False
            
            # Clear fields before entering credentials (in case there's any text)
            username_field.clear()
            password_field.clear()
            
            # Enter credentials
            print(f"Entering username: {username[:3]}*****")
            username_field.send_keys(username)
            time.sleep(1)  # Small delay between fields
            print("Entering password: *****")
            password_field.send_keys(password)
            
            # Take screenshot before clicking login
            self.take_screenshot("credentials_entered")
            
            # Click login button - try multiple possible selectors
            print("Looking for login button...")
            try:
                login_button = self.driver.find_element(By.XPATH, "//button[contains(@class, 'Log') or contains(text(), 'Log In')]")
            except:
                # Try alternate selectors
                login_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[@type='submit'] | //input[@type='submit'] | //button[contains(text(), 'Log')] | //button[contains(text(), 'Sign')]"))
                )
                
            print("Clicking login button...")
            login_button.click()
            
            # Wait for login to complete
            print("Waiting for login to complete...")
            time.sleep(5)
            
            # Check if login was successful
            if self.is_logged_in():
                print("Login successful")
                return True
            print("Login failed - could not detect logged-in state")
            self.take_screenshot("login_failed")
            return False
            
        except Exception as e:
            print(f"Error during login: {e}")
            self.take_screenshot("login_error")
            return False
    
    def run(self):
        """Main method to run the scraping process."""
        try:
            print("\nStarting WeatherLink scraping process...")
            
            # Check if already logged in, if not, log in
            if not self.is_logged_in():
                if not self.login():
                    print("Failed to log in")
                    return None
                print("Login successful")
                time.sleep(3)  # Wait for login to complete
            else:
                print("Already logged in")
            
            # Take a screenshot after login
            self.take_screenshot("after_login")
            
            # Navigate to data page
            if not self.navigate_to_data_page():
                print("Failed to navigate to data page")
                # Don't return None, try to continue
            
            # Set date to current - this step can be skipped if it fails
            print("\n" + "="*50)
            print("ATTEMPTING TO SET DATE TO CURRENT SYSTEM DATE")
            print("="*50)
            date_set_success = self.set_date_to_current()
            if not date_set_success:
                print("\nWARNING: Failed to set date to current system date")
                print("Will proceed with export using the site's default date selection")
                print("The exported data may not reflect the current date")
                # Take a screenshot of the current page state
                self.take_screenshot("date_setting_failed")
            else:
                print("\nDate successfully set to current system date")
                print("Proceeding with data export...")
            
            # Export data to email
            if not self.export_data():
                print("Failed to export data")
                return None
            
            print("Data export request sent successfully")
            
            # If Gmail API is available and we want to check email
            if self.use_api and GmailAPI is not None:
                print("\nWaiting for email with exported data...")
                result = self.check_gmail_api(max_wait_time=300)  # 5 minutes max wait time
                
                # Handle the result, which could be None, filepath, or (filepath, dataframe)
                if result is None:
                    print("Failed to download data from Gmail")
                    return None
                
                # Check if result is a tuple (filepath, dataframe) or just filepath
                if isinstance(result, tuple) and len(result) == 2:
                    csv_file, df = result
                    print(f"Successfully downloaded and processed data from Gmail: {csv_file}")
                    
                    # Process the CSV file - df is already processed, so just save it
                    try:
                        # Save a copy in the current directory with timestamp
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        local_filename = f"weather_data_{timestamp}.csv"
                        df.to_csv(local_filename, index=False)
                        print(f"Saved a copy to {local_filename}")
                        
                        # Process and then delete the temporary CSV file
                        try:
                            # If dataset.csv already exists, we can safely delete the temp file
                            if os.path.exists('dataset.csv'):
                                # Delete the temporary weather_data CSV file
                                import glob
                                for temp_file in glob.glob("weather_data_*.csv"):
                                    try:
                                        os.remove(temp_file)
                                        print(f"Deleted temporary file: {temp_file}")
                                    except Exception as delete_error:
                                        print(f"Warning: Unable to delete {temp_file}: {delete_error}")
                        except Exception as cleanup_error:
                            print(f"Warning: Error during file cleanup: {cleanup_error}")
                            
                        return df
                    except Exception as e:
                        print(f"Error saving CSV copy: {e}")
                        return df
                else:
                    # Got just a filepath, need to process the file
                    csv_file = result
                    print(f"Successfully downloaded data from Gmail: {csv_file}")
                    # Process the downloaded CSV - save original and then clean
                    try:
                        # First save the original file
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        local_filename = f"weather_data_{timestamp}.csv"
                        
                        # Copy the original file
                        import shutil
                        shutil.copy2(csv_file, local_filename)
                        print(f"Saved a copy of the original file to {local_filename}")
                        
                        # Clean the data and save as dataset.csv
                        print("\nCleaning the data...")
                        import pandas as pd
                        
                        # Try different approaches for reading the file
                        try:
                            # Approach 1: Skip the first 5 rows directly during loading
                            print("Attempting to read CSV, skipping first 5 rows...")
                            df = pd.read_csv(csv_file, skiprows=5, encoding='utf-8', 
                                            on_bad_lines='skip', low_memory=False)
                        except UnicodeDecodeError:
                            # If UTF-8 fails, try with latin-1 encoding
                            print("UTF-8 encoding failed, trying with latin-1 encoding...")
                            df = pd.read_csv(csv_file, skiprows=5, encoding='latin-1', 
                                            on_bad_lines='skip', low_memory=False)
                        except pd.errors.ParserError as pe:
                            print(f"Parser error: {pe}")
                            print("Trying with different delimiter...")
                            # Try with a different delimiter
                            df = pd.read_csv(csv_file, skiprows=5, encoding='utf-8', 
                                            on_bad_lines='skip', low_memory=False, delimiter=',', 
                                            quotechar='"')
                        except Exception as e:
                            # If standard approach fails, try a more manual approach
                            print(f"Standard CSV reading failed: {e}")
                            print("Trying manual file reading approach...")
                            
                            # Read raw lines, skip the first 5, then parse
                            with open(csv_file, 'rb') as f:
                                lines = f.readlines()
                            
                            # Skip first 5 lines, then join the rest
                            data_lines = lines[5:]
                            
                            # Write to a temporary file
                            temp_file = 'temp_clean_data.csv'
                            with open(temp_file, 'wb') as f:
                                f.writelines(data_lines)
                            
                            # Now try to read the clean temp file
                            print("Reading from clean temporary file...")
                            df = pd.read_csv(temp_file, encoding='latin-1', 
                                            on_bad_lines='skip', low_memory=False)
                        
                        # Save the cleaned data
                        print(f"Successfully read CSV data with shape: {df.shape}")
                        
                        # Remove specified columns
                        df = self.clean_csv_data(df)
                        
                        # Save the cleaned data
                        df.to_csv('dataset.csv', index=False)
                        print("Cleaned data saved as dataset.csv")
                        
                        # Run clean.py BEFORE returning
                        success = self.run_clean_script()
                        print(f"Clean script run {'successfully' if success else 'with errors'}")
                        
                        # After the data has been downloaded and saved as dataset.csv
                        
                        # Check if the dataset.csv file exists
                        if os.path.exists("dataset.csv"):
                            print("Running data processing script (clean.py)...")
                            try:
                                # Run clean.py as a separate process
                                subprocess.run(["python", "clean.py"], check=True)
                                print("Data processing completed successfully")
                                
                                # Delete all temporary weather_data_*.csv files
                                print("Cleaning up temporary CSV files...")
                                import glob
                                for temp_file in glob.glob("weather_data_*.csv"):
                                    try:
                                        os.remove(temp_file)
                                        print(f"Deleted temporary file: {temp_file}")
                                    except Exception as delete_error:
                                        print(f"Warning: Unable to delete {temp_file}: {delete_error}")
                                
                                # Also delete temp_clean_data.csv if it exists
                                if os.path.exists('temp_clean_data.csv'):
                                    try:
                                        os.remove('temp_clean_data.csv')
                                        print("Deleted temp_clean_data.csv")
                                    except Exception as delete_error:
                                        print(f"Warning: Unable to delete temp_clean_data.csv: {delete_error}")
                                        
                            except subprocess.CalledProcessError as e:
                                print(f"Error running data processing script: {e}")
                            except Exception as e:
                                print(f"Unexpected error running data processing script: {e}")
                        
                        return df
                    except Exception as e:
                        print(f"Error processing downloaded CSV: {e}")
                        return None
            else:
                print("\nGmail API not available or not enabled.")
                print("The export request has been sent to your email.")
                print("You will need to check your email manually for the WeatherLink export.")
                return True
            
        except Exception as e:
            print(f"Error in run method: {e}")
            self.take_screenshot("run_error")
            return None
        finally:
            # Always close the browser
            try:
                self.driver.quit()
                print("Browser closed")
            except:
                pass

    def test_email_extraction(self):
        """Test email extraction functionality"""
        try:
            print("Initializing GmailAPI...")
            gmail_api = GmailAPI()
            if not gmail_api or not hasattr(gmail_api, 'service'):
                print("Error: GmailAPI not properly initialized")
                return False
            
            print("GmailAPI initialized successfully")
            print("Fetching messages...")
            
            try:
                results = gmail_api.service.users().messages().list(
                    userId='me',
                    q='from:weatherlink.com',
                    maxResults=5
                ).execute()
                
                messages = results.get('messages', [])
                if not messages:
                    print("No messages found from WeatherLink")
                    return False
                
                print(f"Found {len(messages)} messages")
                
                for msg in messages:
                    try:
                        message = gmail_api.service.users().messages().get(
                            userId='me',
                            id=msg['id']
                        ).execute()
                        
                        print("\nMessage Details:")
                        print(f"Message ID: {msg['id']}")
                        print(f"Thread ID: {msg['threadId']}")
                        
                        headers = message['payload']['headers']
                        for header in headers:
                            if header['name'].lower() in ['subject', 'from', 'date']:
                                print(f"{header['name']}: {header['value']}")
                        
                        s3_url = None
                        if 'parts' in message['payload']:
                            for part in message['payload']['parts']:
                                if part['mimeType'] == 'text/html':
                                    data = part['body']['data']
                                    if data:
                                        import base64
                                        html = base64.urlsafe_b64decode(data).decode('utf-8')
                                        print(f"\nHTML Content Preview: {html[:500]}...")
                                        
                                        # Look for the download link
                                        import re
                                        href_pattern = r'href="(https://s3\.amazonaws\.com/[^"]+\.csv)"'
                                        match = re.search(href_pattern, html)
                                        if match:
                                            s3_url = match.group(1)
                                            print(f"Found S3 URL in HTML: {s3_url}")
                                            break
                                        
                                        # Try alternative pattern
                                        s3_pattern = r'https://s3\.amazonaws\.com/export-wl2-live\.weatherlink\.com/data/[^"]+\.csv'
                                        match = re.search(s3_pattern, html)
                                        if match:
                                            s3_url = match.group(0)
                                            print(f"Found S3 URL: {s3_url}")
                                            break
                                        
                                        print("No S3 URL patterns matched in HTML content")
                                else:
                                    print("Part has no data content")
                        else:
                            print("Email doesn't have the expected 'parts' structure")
                            if 'body' in message['payload'] and 'data' in message['payload']['body']:
                                data = message['payload']['body']['data']
                                if data:
                                    import base64
                                    html = base64.urlsafe_b64decode(data).decode('utf-8')
                                    print(f"Direct body content preview: {html[:200]}...")
                                    
                                    # Look for the download link
                                    import re
                                    href_pattern = r'href="(https://s3\.amazonaws\.com/[^"]+\.csv)"'
                                    match = re.search(href_pattern, html)
                                    if match:
                                        s3_url = match.group(1)
                                        print(f"Found S3 URL in body: {s3_url}")
                        
                        # Save a copy of the email for debugging
                        debug_dir = 'debug_emails'
                        if not os.path.exists(debug_dir):
                            os.makedirs(debug_dir)
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        debug_file = f"{debug_dir}/email_debug_{timestamp}.txt"
                        with open(debug_file, 'w', encoding='utf-8') as f:
                            f.write(f"Message ID: {msg['id']}\n\n")
                            f.write("Headers:\n")
                            for header in message['payload']['headers']:
                                if header['name'].lower() in ['subject', 'from', 'date']:
                                    f.write(f"{header['name']}: {header['value']}\n")
                            f.write("\nFull Message Structure:\n")
                            import json
                            f.write(json.dumps(message, indent=2))
                        print(f"Saved debug email information to {debug_file}")
                        
                        if s3_url:
                            print(f"\nFound S3 URL: {s3_url}")
                            # Create directory for saved files if it doesn't exist
                            if not os.path.exists('saved_emails'):
                                os.makedirs('saved_emails')
                            
                            # Download the file
                            import urllib.request
                            filename = s3_url.split('/')[-1]
                            filepath = os.path.join('saved_emails', filename)
                            
                            print(f"Downloading file to {filepath}...")
                            urllib.request.urlretrieve(s3_url, filepath)
                            print("Download complete!")
                            
                            # Clean the data and save as dataset.csv
                            print("\nCleaning the data...")
                            import pandas as pd
                            try:
                                # Try different approaches for reading the file
                                try:
                                    # Approach 1: Skip the first 5 rows directly during loading
                                    print("Attempting to read CSV, skipping first 5 rows...")
                                    df = pd.read_csv(filepath, skiprows=5, encoding='utf-8', 
                                                    on_bad_lines='skip', low_memory=False)
                                except UnicodeDecodeError:
                                    # If UTF-8 fails, try with latin-1 encoding
                                    print("UTF-8 encoding failed, trying with latin-1 encoding...")
                                    df = pd.read_csv(filepath, skiprows=5, encoding='latin-1', 
                                                    on_bad_lines='skip', low_memory=False)
                                except pd.errors.ParserError as pe:
                                    print(f"Parser error: {pe}")
                                    print("Trying with different delimiter...")
                                    # Try with a different delimiter
                                    df = pd.read_csv(filepath, skiprows=5, encoding='utf-8', 
                                                    on_bad_lines='skip', low_memory=False, delimiter=',', 
                                                    quotechar='"')
                                except Exception as e:
                                    # If standard approach fails, try a more manual approach
                                    print(f"Standard CSV reading failed: {e}")
                                    print("Trying manual file reading approach...")
                                    
                                    # Read raw lines, skip the first 5, then parse
                                    with open(filepath, 'rb') as f:
                                        lines = f.readlines()
                                    
                                    # Skip first 5 lines, then join the rest
                                    data_lines = lines[5:]
                                    
                                    # Write to a temporary file
                                    temp_file = 'temp_clean_data.csv'
                                    with open(temp_file, 'wb') as f:
                                        f.writelines(data_lines)
                                    
                                    # Now try to read the clean temp file
                                    print("Reading from clean temporary file...")
                                    df = pd.read_csv(temp_file, encoding='latin-1', 
                                                    on_bad_lines='skip', low_memory=False)
                                
                                # Save the cleaned data
                                print(f"Successfully read CSV data with shape: {df.shape}")
                                
                                # Remove specified columns
                                df = self.clean_csv_data(df)
                                
                                # Save the clean data to dataset.csv
                                df.to_csv('dataset.csv', index=False)
                                print("Cleaned data saved as dataset.csv")
                                
                                # Run clean.py BEFORE returning
                                success = self.run_clean_script()
                                print(f"Clean script run {'successfully' if success else 'with errors'}")
                                
                                # Now return the dataframe
                                return df
                            except Exception as csv_error:
                                print(f"Error processing CSV file: {csv_error}")
                                # Continue to next message if this one fails
                                continue
                        
                    except Exception as get_error:
                        print(f"Error getting message details: {get_error}")
                        continue
                
                # If we get here, we didn't find any S3 URLs in any messages
                print("\nNo S3 URLs found in any messages")
                return False
                
            except Exception as list_error:
                print(f"Error listing messages: {list_error}")
                if 'insufficient authentication scopes' in str(list_error).lower():
                    print("\nGmail API access issue detected. Please ensure you have:")
                    print("1. Enabled the Gmail API in Google Cloud Console")
                    print("2. Created OAuth 2.0 credentials")
                    print("3. Downloaded the credentials.json file")
                    print("4. Run the script to authenticate")
                return False
            
        except Exception as e:
            print(f"Error in test_email_extraction: {e}")
            return False

    def download_csv_from_email(self):
        """Download the CSV file from the S3 URL found in the email."""
        try:
            print("\nDownloading CSV file from email...")
            
            if GmailAPI is None:
                print("GmailAPI module is not available")
                return None
            
            # Initialize GmailAPI
            gmail_api = GmailAPI()
            if not gmail_api or not gmail_api.service:
                print("Failed to initialize GmailAPI")
                return None
            
            # Get the latest message from WeatherLink
            messages = gmail_api.service.users().messages().list(
                userId='me',
                q='from:weatherlink.com',
                maxResults=1
            ).execute()
            
            if not messages.get('messages'):
                print("No messages found from WeatherLink")
                return None
            
            # Get the message details
            msg = gmail_api.service.users().messages().get(
                userId='me',
                id=messages['messages'][0]['id'],
                format='full'
            ).execute()
            
            # Extract the S3 URL from the email content
            s3_url = None
            if 'parts' in msg['payload']:
                for part in msg['payload']['parts']:
                    if part['mimeType'] == 'text/html':
                        data = part['body'].get('data', '')
                        if data:
                            import base64
                            html = base64.urlsafe_b64decode(data).decode('utf-8')
                            # Look for the download link in the HTML
                            import re
                            s3_url_pattern = r'https://s3\.amazonaws\.com/export-wl2-live\.weatherlink\.com/data/[^"]+\.csv'
                            match = re.search(s3_url_pattern, html)
                            if match:
                                s3_url = match.group(0)
                                print(f"Found S3 URL: {s3_url}")
                                break
            
            if not s3_url:
                print("No S3 URL found in the email")
                return None
            
            # Create a directory for saved files if it doesn't exist
            if not os.path.exists('saved_emails'):
                os.makedirs('saved_emails')
            
            # Extract filename from URL
            filename = s3_url.split('/')[-1]
            filepath = f"saved_emails/{filename}"
            
            # Download the file directly using requests
            print(f"Downloading file from {s3_url}...")
            try:
                import requests
                response = requests.get(s3_url)
                if response.status_code == 200:
                    with open(filepath, 'wb') as f:
                        f.write(response.content)
                    print(f"Successfully downloaded file: {filepath}")
                    return filepath
                else:
                    print(f"Failed to download file. Status code: {response.status_code}")
                    return None
            except Exception as download_error:
                print(f"Error downloading file: {download_error}")
                return None
            
        except Exception as e:
            print(f"Error downloading CSV file: {e}")
            return None

    def excel_col_to_index(self, col_str):
        """Convert Excel column letter to 0-based index"""
        result = 0
        for c in col_str:
            result = result * 26 + (ord(c.upper()) - ord('A') + 1)
        return result - 1  # Convert to 0-based index

    def clean_csv_data(self, df):
        """Clean the dataframe by removing specified columns"""
        print("Removing specified columns...")
        
        # Columns to remove by Excel notation
        cols_to_remove = ['D','E','G','H','J','K','M','N','P','Q','R','S','T','U','V','W','X','Y','AA','AB','AC','AD','AE','AG','AH','AJ','AK','AL','AM']
        
        # Convert column letters to numeric indices
        indices_to_remove = []
        for col in cols_to_remove:
            idx = self.excel_col_to_index(col)
            if idx < len(df.columns):
                indices_to_remove.append(idx)
        
        # Get actual column names to drop using positional indices
        cols_to_drop = [df.columns[i] for i in indices_to_remove if i < len(df.columns)]
        
        if cols_to_drop:
            print(f"Dropping {len(cols_to_drop)} columns")
            df = df.drop(columns=cols_to_drop)
            print(f"DataFrame shape after removing columns: {df.shape}")
        else:
            print("No matching columns found to remove")
        
        return df
    
    def process_hourly_averages(self, df):
        """
        Process data to create hourly averages from 5-minute increments.
        Only include hours with complete data (all 12 5-minute increments from :00 to :55).
        Separate Date & Time into Date, Start Period, and End Period columns.
        
        Args:
            df: DataFrame with cleaned data
            
        Returns:
            DataFrame with hourly averaged data
        """
        print("Processing hourly averages from 5-minute increment data...")
        
        # Identify the Date & Time column (usually the first column)
        datetime_col = df.columns[0]
        print(f"Using '{datetime_col}' as the date/time column")
        
        try:
            # Convert the Date & Time column to datetime format
            df['datetime'] = pd.to_datetime(df[datetime_col], format='%m/%d/%Y %I:%M:%S %p')
            print("Converted date column to datetime format")
            
            # Create date and hour components
            df['date'] = df['datetime'].dt.date
            df['hour'] = df['datetime'].dt.hour
            df['minute'] = df['datetime'].dt.minute
            
            # Group by date and hour to count how many 5-minute increments exist for each hour
            count_by_hour = df.groupby(['date', 'hour'])['minute'].count()
            
            # Filter to keep only those date-hour combinations that have all 12 increments (0, 5, 10, ..., 55)
            complete_hours = count_by_hour[count_by_hour == 12].reset_index()[['date', 'hour']]
            
            # Merge with the original data to filter out incomplete hours
            df = pd.merge(df, complete_hours, on=['date', 'hour'])
            print(f"Filtered to keep only complete hours with all 12 5-minute increments: {len(df)} rows remaining")
            
            # Calculate hourly averages
            # Exclude the datetime, date, hour, and minute columns from averaging
            numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
            numeric_cols = [col for col in numeric_cols if col not in ['hour', 'minute']]
            
            # Group by date and hour and calculate averages
            hourly_avg = df.groupby(['date', 'hour'])[numeric_cols].mean().reset_index()
            
            # Add Start Period and End Period columns
            hourly_avg['Start Period'] = hourly_avg.apply(
                lambda x: f"{x['hour']}:00 {'AM' if x['hour'] < 12 else 'PM'}", axis=1)
            hourly_avg['End Period'] = hourly_avg.apply(
                lambda x: f"{x['hour']}:55 {'AM' if x['hour'] < 12 else 'PM'}", axis=1)
            
            # Format the Start Period and End Period using 12-hour clock
            hourly_avg['Start Period'] = hourly_avg.apply(
                lambda x: f"{x['hour'] if x['hour'] <= 12 else x['hour'] - 12}:00 {'AM' if x['hour'] < 12 else 'PM'}", 
                axis=1)
            hourly_avg['End Period'] = hourly_avg.apply(
                lambda x: f"{x['hour'] if x['hour'] <= 12 else x['hour'] - 12}:55 {'AM' if x['hour'] < 12 else 'PM'}", 
                axis=1)
            
            # Handle 0 hour (12 AM)
            hourly_avg.loc[hourly_avg['hour'] == 0, 'Start Period'] = "12:00 AM"
            hourly_avg.loc[hourly_avg['hour'] == 0, 'End Period'] = "12:55 AM"
            
            # Handle 12 hour (12 PM)
            hourly_avg.loc[hourly_avg['hour'] == 12, 'Start Period'] = "12:00 PM"
            hourly_avg.loc[hourly_avg['hour'] == 12, 'End Period'] = "12:55 PM"
            
            # Format the Date column
            hourly_avg['Date'] = hourly_avg['date'].astype(str)
            
            # Final column selection and reordering
            final_columns = ['Date', 'Start Period', 'End Period'] + [col for col in hourly_avg.columns 
                                                                     if col not in ['date', 'hour', 'Date', 'Start Period', 'End Period']]
            final_df = hourly_avg[final_columns]
            
            print(f"Successfully created hourly averages. Final shape: {final_df.shape}")
            return final_df
            
        except Exception as e:
            print(f"Error processing hourly averages: {e}")
            import traceback
            traceback.print_exc()
            return df  # Return original dataframe if processing fails

    def run_clean_script(self):
        """Run the clean.py script after dataset.csv has been created"""
        print("=" * 50)
        print("RUNNING DATA PROCESSING SCRIPT (clean.py)")
        print("=" * 50)
        try:
            # Check if clean.py exists
            if not os.path.exists("clean.py"):
                print("ERROR: clean.py not found in current directory!")
                return False
                
            # Run clean.py as a separate process
            print("Executing: python clean.py")
            result = subprocess.run(["python", "clean.py"], 
                                   check=True, 
                                   capture_output=True, 
                                   text=True)
            
            # Print the output from clean.py
            if result.stdout:
                print("Output from clean.py:")
                print(result.stdout)
                
            if result.stderr:
                print("Errors from clean.py:")
                print(result.stderr)
                
            print("Data processing completed successfully")
            return True
        except subprocess.CalledProcessError as e:
            print(f"Error running data processing script: {e}")
            if e.stdout:
                print("Output:")
                print(e.stdout)
            if e.stderr:
                print("Error output:")
                print(e.stderr)
            return False
        except Exception as e:
            print(f"Unexpected error running data processing script: {e}")
            return False

# Add main execution block
if __name__ == "__main__":
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='WeatherLink data scraper')
    parser.add_argument('--url', type=str, default='https://www.weatherlink.com/',
                      help='WeatherLink URL (default: https://www.weatherlink.com/)')
    parser.add_argument('--debug', action='store_true',
                      help='Enable debug mode (takes screenshots)')
    parser.add_argument('--export-email', type=str, default='teamdavcast@gmail.com',
                      help='Email address to export data to (default: teamdavcast@gmail.com)')
    parser.add_argument('--check-email', action='store_true',
                      help='Check email for exported data')
    parser.add_argument('--use-api', action='store_true',
                      help='Use Gmail API for email checking')
    parser.add_argument('--no-api', action='store_true',
                      help='Force using browser automation instead of API')
    parser.add_argument('--wait-time', type=int, default=300,
                      help='Maximum time to wait for email (in seconds)')
    parser.add_argument('--test-email', action='store_true',
                      help='Test Gmail API email extraction')
    parser.add_argument('--download-csv', action='store_true',
                      help='Download the CSV file from the latest WeatherLink email')
    
    args = parser.parse_args()
    
    # Load environment variables
    load_dotenv()
    
    # Get WeatherLink URL from environment or use default
    weatherlink_url = os.getenv('WEATHERLINK_URL', args.url)
    
    print("\n" + "="*50)
    print("WEATHERLINK DATA SCRAPER")
    print("="*50)
    print(f"URL: {weatherlink_url}")
    print(f"Debug mode: {'ON' if args.debug else 'OFF'}")
    print(f"Export email: {args.export_email}")
    print(f"Check Gmail: {'YES' if args.check_email else 'NO'}")
    print(f"Use Gmail API: {'YES' if args.use_api and GmailAPI is not None and not args.no_api else 'NO'}")
    print("="*50 + "\n")
    
    try:
        # If test-email flag is set, run the test without initializing browser
        if args.test_email:
            print("\nRunning Gmail API email extraction test...")
            # Create a minimal instance without browser initialization
            scraper = WeatherLink(args.url, debug=False, export_email=args.export_email, init_browser=False)
            if scraper.test_email_extraction():
                print("\nEmail extraction test completed successfully!")
            else:
                print("\nEmail extraction test failed. Check the error messages above.")
            sys.exit(0)
        
        # If download-csv flag is set, download the CSV file
        if args.download_csv:
            print("\nDownloading CSV file from latest WeatherLink email...")
            # Create a minimal instance without browser initialization
            scraper = WeatherLink(args.url, debug=False, export_email=args.export_email, init_browser=False)
            csv_file = scraper.download_csv_from_email()
            if csv_file:
                print(f"\nSuccessfully downloaded CSV file: {csv_file}")
            else:
                print("\nFailed to download CSV file. Check the error messages above.")
            sys.exit(0)
        
        # For all other operations, create a full instance with browser
        scraper = WeatherLink(weatherlink_url, debug=args.debug, export_email=args.export_email)
        
        # Set API usage based on command line arguments
        if args.no_api:
            scraper.use_api = False
        elif args.use_api and GmailAPI is not None:
            scraper.use_api = True
        
        # Run the scraping process
        print("Beginning scraping process...")
        result = scraper.run()
        
        # If we need to check email but run() didn't do it already
        if args.check_email and result is True:  # Export succeeded but no data downloaded yet
            print("\nExplicitly checking Gmail as requested...")
            
            # Use API if requested and available
            csv_file = None
            if args.use_api and GmailAPI is not None and not args.no_api:
                print("Using Gmail API to check for the exported data...")
                csv_file = scraper.check_gmail_api(max_wait_time=args.wait_time)
            
            # Fall back to browser automation if API fails or not available
            if csv_file is None or args.no_api:
                print("Using browser automation to check Gmail...")
                # TODO: Implement browser automation for Gmail if needed
                print("Browser automation for Gmail not implemented yet")
            
            if csv_file:
                print(f"\nSuccessfully downloaded data from Gmail: {csv_file}")
                # Process the CSV file
                try:
                    import pandas as pd
                    df = pd.read_csv(csv_file)
                    # Process hourly averages
                    df = scraper.process_hourly_averages(df)
                    # Save a copy in the current directory with timestamp
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    local_filename = f"weather_data_{timestamp}.csv"
                    df.to_csv(local_filename, index=False)
                    print(f"Saved a copy to {local_filename}")
                    # Save the hourly averaged data to dataset.csv
                    df.to_csv('dataset.csv', index=False)
                    print("Saved hourly averaged data to dataset.csv")
                    result = df  # Update result to the DataFrame
                except Exception as e:
                    print(f"Error processing downloaded CSV: {e}")
        
        if result is not None:
            print("\n" + "="*50)
            print("SCRAPING COMPLETED SUCCESSFULLY!")
            if os.path.exists('dataset.csv'):
                print(f"Original data saved to weather_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
                print(f"Cleaned data saved to dataset.csv")
            else:
                print(f"Data saved to weather_data.csv")
            print("="*50 + "\n")
        else:
            print("\n" + "="*50)
            print("SCRAPING PROCESS ENCOUNTERED ISSUES")
            print("Check the error messages above and the screenshots folder")
            print("="*50 + "\n")
    except Exception as e:
        print(f"\nError in main execution: {e}")
        print("Please check your credentials, internet connection, and Chrome installation.")
