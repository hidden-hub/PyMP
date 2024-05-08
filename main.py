import requests
import openpyxl
import re
from bs4 import BeautifulSoup
import csv

def get_links_from_excel():
    
    links = []
    
    wb = openpyxl.load_workbook("links.xlsx")
    
    sheet = wb['Лист1']
    
    for row in sheet.iter_rows(min_row=2, values_only=True):
        for cell in row:
            if isinstance(cell, str) and "http" in cell:
                links.append(cell)
                print(f'prased link {cell}')
            else:
                print(f'cell {cell} is not a string or does not contain a link')
    return links


def parse_from_link():

    links = get_links_from_excel()
    
    all_emails = []
    
    for site in links:
        try:
            response = requests.get(site)
            response.raise_for_status()  # Raises an HTTPError for bad responses
            soup = BeautifulSoup(response.text, 'html.parser')
            emails = set(re.findall(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", soup.get_text()))
            all_emails.extend(emails)
            print(f'parsed from {site} {emails}')
        except (requests.exceptions.RequestException, requests.exceptions.HTTPError) as e:
            print(f'Failed to parse or connect to {site}. Error: {e}')
    return all_emails


if __name__ == "__main__":

    emails = parse_from_link()
    
    with open('emails.csv', 'w', newline='') as file:
        writer = csv.writer(file)
        for email in emails:
            print(f'wrote email {email}')
            writer.writerow([email])
