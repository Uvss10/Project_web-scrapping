import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup

# Setup
INPUT_FILE = 'Input.xlsx'
OUTPUT_FOLDER = 'text_files'
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
df = pd.read_excel(INPUT_FILE)
urls = df[['URL_ID', 'URL']].values.tolist()

# Chrome options
options = Options()
options.add_argument("--headless=new")
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome(options=options)

def scrape_with_selenium(url):
    try:
        driver.get(url)
        time.sleep(3)
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # Title
        title_tag = soup.find('h1', class_='entry-title')
        if not title_tag:
            for h1 in soup.find_all('h1'):
                if h1.get_text(strip=True):
                    title_tag = h1
                    break
        title = title_tag.get_text(strip=True) if title_tag else 'No Title Found'

        # Content
        content_div = soup.find('div', class_='td-post-content')
        content_parts = []

        if content_div:
            # Remove unwanted tags
            for tag in content_div(['script', 'style', 'aside', 'footer']):
                tag.decompose()

            # Iterate through children to maintain structure
            for elem in content_div.find_all(recursive=False):
                text = elem.get_text(separator=' ', strip=True)

                if not text:
                    continue

                # Normalize section formatting
                if "Project Description" in text:
                    content_parts.append("\nProject Description\n")
                elif "Data Visualization Deliverables" in text:
                    content_parts.append("\nData Visualization Deliverables\n")
                elif "Data Visualization Tools" in text:
                    content_parts.append("\nData Visualization Tools\n")
                elif "Data Visualization Languages" in text:
                    content_parts.append("\nData Visualization Languages\n")
                elif "Demo" in text:
                    content_parts.append("\nDemo\n")
                else:
                    content_parts.append(text)

        content = '\n'.join(content_parts) if content_parts else 'No Content Found'
        return title, content

    except Exception as e:
        print(f"[Exception] {url}: {e}")
        return None, None

# Run
for url_id, url in urls:
    title, content = scrape_with_selenium(url)
    if title and content:
        filepath = os.path.join(OUTPUT_FOLDER, f"{url_id}.txt")
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f"{title}\n\n{content}")
        print(f"[✓] Saved: {filepath}")
    else:
        print(f"[✗] Skipped: {url}")

driver.quit()
