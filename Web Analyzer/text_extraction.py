#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import requests
from bs4 import BeautifulSoup
import os

import requests
from bs4 import BeautifulSoup

def scrape10(url,url_id):
        try:
            response = requests.get(url)
            if response.status_code == 200:
              soup = BeautifulSoup(response.text, 'html.parser')
              article_name = soup.find("h1",class_='tdb-title-text')
              if article_name:
                article_name = article_name.text.strip()
              article_text = soup.find("div",class_='td_block_wrap tdb_single_content tdi_130 td-pb-border-top td_block_template_1 td-post-content tagdiv-type')
              if article_text:
                article_text = article_text.text.strip()
              return  article_name,article_text
            else:
             file_name = f"{url_id}.txt"
             with open(file_name, 'w', encoding='utf-8') as file:
                file.write("NA")
        except Exception as e:
            print(f"Error scraping {url}: {e}")
            return None, None

def scrape_urls_from_excel(excel_file):
    df = pd.read_excel(excel_file)
    for index, row in df.iterrows():
        url = row['URL']
        url_id = row['URL_ID']
        try:
            response = requests.get(url)
            soup = BeautifulSoup(response.text, 'html.parser')
            article_name = soup.find("h1", class_='entry-title')
            if article_name:
                article_name = article_name.text.strip()
                article_text = soup.find("div", class_='td-post-content tagdiv-type')
                if article_text:
                    article_text = article_text.text.strip()
                else:
                    raise ValueError("Article text not found.")
                file_name = f"{url_id}.txt"
                with open(file_name, 'w', encoding='utf-8') as file:
                    file.write(f"Article Name: {article_name}\n")
                    file.write(f"Article Text: {article_text}")
                print(f"Article '{article_name}' saved to '{file_name}'")
            else:
                article_name, article_text = scrape10(url,url_id)
                if article_name and article_text:
                    file_name = f"{url_id}.txt"
                    with open(file_name, 'w', encoding='utf-8') as file:
                        file.write(f"Article Name: {article_name}\n")
                        file.write(f"Article Text: {article_text}")
                    print(f"Article '{article_name}' saved to '{file_name}'")
            print()
        except Exception as e:
            print(f"Error processing URL {url}: {e}")

def main():
    excel_file = r'C:\\Users\\gayat\\OneDrive\\Desktop\\BlackCoffer\\Input.xlsx'
    scrape_urls_from_excel(excel_file)

if __name__ == "__main__":
    main()


# In[ ]:




