import requests
from bs4 import BeautifulSoup
import json
import pandas as pd

def fetch_data():
    url = "https://merolagani.com/LatestMarket.aspx"
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'lxml')
    table = soup.find('table', class_= 'table table-hover live-trading sortable')
    headers = [header.text for header in table.find_all('th')]
    rows = table.find_all('tr',{"class": ['decrease=row', 'increase-row', 'nochange-row']})
    result = [{headers[index]: cell.text for index, cell in enumerate(row.find_all('td'))} for row in rows]
    fullname = [tag.get('title') for tag in table.find_all('a', {"target": ["_blank"]})]
    data = [{'Name': fullname[i], 'Symbol': result[i]['Symbol'], 'LTP': result[i]['LTP'], 
             'Change': result[i]['% Change'], 'High': result[i]['High'], 'Low': result[i]['Low'], 
             'Open': result[i]['Open'], 'Qty.': result[i]['Qty.']} for i in range(len(result))]
    return data

def main():
    data = fetch_data()
    df = pd.DataFrame(data)
    df.to_excel("sharedata.xlsx")
    for item in data:
        print(json.dumps(item, indent=2))

if __name__ == "__main__":
    main()