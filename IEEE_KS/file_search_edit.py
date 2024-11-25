import os
import time
import logging
import pandas as pd
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configure logging to display messages
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Define the folder containing your CSV files
folder_path = #enter your path here

# Define the output path for the results file
output_file_path = # enter your path here

# Keywords to search for in the article's full text
meta_keyword = 'meta'
keywords_to_check = ['qual', 'qualsyst']

# Path to the Edge WebDriver executable
edge_driver_path = # enter your path here

# Ensure the driver executable exists at the specified path
if not os.path.exists(edge_driver_path):
    logger.error(f"Edge WebDriver not found at {edge_driver_path}")
    exit(1)  # Exit the script if the driver is not found

# Selenium setup with Edge options
edge_options = Options()
edge_options.add_argument('--ignore-certificate-errors')
# Suppress logging messages from EdgeDriver
edge_options.add_argument('--log-level=3')
# Uncomment the next line to run Edge in headless mode
# edge_options.add_argument('--headless')

# Initialize Edge WebDriver
driver = webdriver.Edge(service=Service(edge_driver_path), options=edge_options)

# Function to normalize ProQuest URLs
def normalize_proquest_url(csv_url):
    base_url = "https://www.proquest.com/docview/"
    try:
        # Extract the identifier after '/docview/' and before '/abstract/'
        docview_id = csv_url.split('/docview/')[1].split('/abstract/')[0]
        # Extract the session-related identifier and append it
        session_id = csv_url.split('/abstract/')[1].split('?')[0]
        # Recreate the normalized URL
        normalized_url = f"{base_url}{docview_id}/{session_id}?accountid=14771"
        return normalized_url
    except IndexError:
        # Handle case where the URL does not match expected format
        logger.error(f"Error processing URL: {csv_url}")
        return csv_url  # Return the original URL if there's an issue

# Function to dismiss cookie consent
def dismiss_cookie_consent():
    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler"))
        ).click()
        logger.info("Cookie consent dismissed")
    except Exception as e:
        logger.info("No cookie consent pop-up detected or failed to close.")

# Function to wait for overlay to disappear
def wait_for_overlay_to_disappear():
    try:
        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, "onetrust-pc-dark-filter"))
        )
        logger.info("Overlay has disappeared.")
    except Exception as e:
        logger.info("Overlay did not disappear.")

# Function to keep the WebDriver session alive during login
def keep_alive_during_login(wait_time):
    start_time = time.time()
    while time.time() - start_time < wait_time:
        try:
            # Perform a simple JavaScript execution to keep the session alive
            driver.execute_script("return document.title;")
            time.sleep(15)  # Wait for 15 seconds before the next interaction
        except Exception as e:
            logger.error(f"Keep-alive failed: {e}")
            break

# Log into the UofT portal through ProQuest
def login_to_uoft():
    try:
        driver.get("https://www.proquest.com")
        logger.info("Navigated to ProQuest")

        dismiss_cookie_consent()
        wait_for_overlay_to_disappear()

        # Proceed to click "Log in through your library"
        WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "createLoginOverlay"))
        ).click()
        logger.info("Login process initiated. Please log in manually.")

        # Wait for 3 minutes with keep-alive
        login_wait_time = 180  # 180 seconds = 3 minutes
        logger.info(f"Waiting for {login_wait_time // 60} minutes for you to log in...")
        keep_alive_during_login(login_wait_time)
        logger.info("Proceeding with the script...")
    except Exception as e:
        logger.error(f"Error during login: {e}")
        # Do not quit the driver here to avoid closing the browser prematurely
        pass

# Function to wait for CAPTCHA resolution
def wait_for_captcha_resolution():
    logger.info("CAPTCHA detected. Please resolve it manually in the browser.")
    input("Once you have resolved the CAPTCHA, press Enter to continue...")

# Function to fetch full-text content of an article
def fetch_full_text(url):
    global driver  # Declare driver as global to ensure we use the same instance
    try:
        # Check if the driver session is still active
        if driver.session_id is None:
            logger.warning("WebDriver session has ended. Restarting the driver.")
            driver.quit()
            driver = webdriver.Edge(service=Service(edge_driver_path), options=edge_options)
            login_to_uoft()  # Re-login if necessary

        driver.get(url)
        
        # Check for CAPTCHA and wait for resolution
        if "robot" in driver.page_source.lower():
            wait_for_captcha_resolution()
        
        # Wait for the body of the page to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        text = driver.find_element(By.TAG_NAME, "body").text.lower()
        return text
    except Exception as e:
        logger.error(f"Error fetching full text for {url}: {e}")
        return "error"  # Return "error" for unsuccessful fetch

# Initialize results list
meta_analysis_results = []

# Login once to maintain session
login_to_uoft()

# Process each file in the folder
for filename in os.listdir(folder_path):
    if filename.endswith('.csv'):
        file_path = os.path.join(folder_path, filename)
        logger.info(f"Processing file: {filename}")
        try:
            df = pd.read_csv(file_path)

            # Check for required columns
            if 'Article Title' in df.columns and 'Article Link' in df.columns:
                matching_articles = df[df['Article Title'].str.contains(meta_keyword, case=False, na=False)]

                for _, row in matching_articles.iterrows():
                    # Normalize the URL before using it
                    original_link = row['Article Link'].strip().strip('"')
                    formatted_link = normalize_proquest_url(original_link)  # Normalize the URL

                    full_text = fetch_full_text(formatted_link)  # Fetch the full article text
                    has_keywords = False if full_text == "error" else any(keyword in full_text for keyword in keywords_to_check)

                    meta_analysis_results.append({
                        'File': filename,
                        'Title': row['Article Title'],
                        'Original URL': original_link,  # Include the original URL
                        'Formatted URL': formatted_link,  # Include the formatted URL
                        'Has QUALSYST Score': has_keywords if full_text != "error" else "error"  # Handle error cases
                    })
            else:
                logger.warning(f"Missing required columns in {filename}. Skipping this file.")
        except Exception as e:
            logger.error(f"Error processing file {filename}: {e}")

# Close the browser
try:
    driver.quit()
except Exception as e:
    logger.error(f"Error closing the browser: {e}")

# Save results to CSV
try:
    if os.path.exists(output_file_path):
        os.remove(output_file_path)
        logger.info(f"Existing file '{output_file_path}' deleted.")
    results_df = pd.DataFrame(meta_analysis_results)
    results_df.to_csv(output_file_path, index=False)
    logger.info(f"Processing complete. Results saved to '{output_file_path}'.")
except Exception as e:
    logger.error(f"Error saving the results to CSV: {e}")
