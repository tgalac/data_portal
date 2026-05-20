import json
import requests
from bs4 import BeautifulSoup

def scrape_hnb():
    url = "https://www.hnb.hr/statistika/statisticka-priopcenja"

    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, 'html.parser')
        headings = soup.find_all('h4')
        
        results = []
        for i in range(min(len(headings), 3)):
            title = headings[i].get_text(strip=True)
            
            # --- NEW UNIQUE LINK DETECTOR LOGIC ---
            # Try to look for an <a> tag directly inside or around the h4 element
            link_tag = headings[i].find('a')
            if link_tag and link_tag.get('href'):
                article_url = link_tag.get('href')
                # If HNB uses absolute paths like "/-/objava...", stitch the domain back on
                if article_url.startswith('/'):
                    article_url = f"https://www.hnb.hr{article_url}"
            else:
                # Fallback to main overview index link if no specific tag exists
                article_url = url
            
            next_el = headings[i].find_next_sibling()
            date = next_el.get_text(strip=True) if next_el and next_el.name == 'h5' else ""
            
            results.append({
                "title": title,
                "date": date,
                "url": article_url # Now maps dynamically to the target link!
            })
            
        with open("SPlast3.json", "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=4)
            
        print("Successfully updated SPlast3.json with unique target links!")
        
    except Exception as e:
        print(f"Error occurred during scraping: {e}")
        with open("SPlast3.json", "w", encoding="utf-8") as f:
            json.dump([], f)

if __name__ == "__main__":
    scrape_hnb()
