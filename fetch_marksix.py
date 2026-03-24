import requests, pandas as pd
from bs4 import BeautifulSoup

url='https://lottery.hk/en/mark-six/results/'
html=requests.get(url,timeout=20).text
soup=BeautifulSoup(html,'html.parser')
rows=soup.select('table tbody tr')
records=[]
for r in rows:
    cols=[c.text.strip() for c in r.select('td')]
    if len(cols)>=8 and cols[0].isdigit() is False:
        try:
            records.append({
              '期數':cols[0],'日期':cols[1],'N1':int(cols[2]),'N2':int(cols[3]),'N3':int(cols[4]),
              'N4':int(cols[5]),'N5':int(cols[6]),'N6':int(cols[7]),'特別號':int(cols[8])
            })
        except: pass

df=pd.DataFrame(records)[::-1]
df.to_excel('data.xlsx',index=False)
