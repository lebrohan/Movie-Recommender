import requests
from bs4 import BeautifulSoup

USERNAME = "dave"
headers = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"}

# Try the RSS feed — server-rendered, includes ratings
rss = requests.get(f"https://letterboxd.com/{USERNAME}/rss/", headers=headers)
print("=== RSS (first 3000 chars) ===")
print(rss.text[:3000])