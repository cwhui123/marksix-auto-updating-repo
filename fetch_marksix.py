# fetch_marksix.py
# Safe version with debug output

import requests
import pandas as pd
from bs4 import BeautifulSoup

url = "https://lottery.hk/en/mark-six/results/"
html = requests.get(url, timeout=20).text
soup = BeautifulSoup(html, "html.parser")

rows = soup.find_all("tr")
records = []

for r in rows:
    tds = r.find_all("td")
    if len(tds) >= 9:
        try:
            records.append({
                "期數": tds[0].text.strip(),
                "日期": tds[1].text.strip(),
                "N1": int(tds[2].text.strip()),
                "N2": int(tds[3].text.strip()),
                "N3": int(tds[4].text.strip()),
                "N4": int(tds[5].text.strip()),
                "N5": int(tds[6].text.strip()),
                "N6": int(tds[7].text.strip()),
                "特別號": int(tds[8].text.strip()),
            })
        except Exception:
            continue

print(f"✅ Records fetched: {len(records)}")

df = pd.DataFrame(records[::-1])
df.to_excel("data.xlsx", index=False)
print("✅ data.xlsx written")
