import json
import requests
from bs4 import BeautifulSoup

def scrape_hnb():
    url = "https://www.hnb.hr/statistika/statisticka-priopcenja"

    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, 'html.parser')
        headings = soup.find_all('h4') # Finds the articles
        
        results = []
        # Get up to the top 3 elements
        for i in range(min(len(headings), 3)):
            title = headings[i].get_text(strip=True)
            
            # Find the date heading right beneath it (h5 element)
            next_el = headings[i].find_next_sibling()
            date = next_el.get_text(strip=True) if next_el and next_el.name == 'h5' else ""
            
            results.append({
                "title": title,
                "date": date,
                "url": url
            })
            
        # Write results to a static file inside the repo
        with open("data.json", "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=4)
            
        print("Successfully updated data.json!")
        
    except Exception as e:
        print(f"Error occurred during scraping: {e}")

if __name__ == "__main__":
    scrape_hnb()
