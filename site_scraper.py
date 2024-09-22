import requests
from bs4 import BeautifulSoup

class SiteScraper:
    
    def __init__(self, url):
        self.url = url
        self.soup = None
    
    def fetch(self):
        try:
            response = requests.get(self.url)
            response.raise_for_status()
            self.soup = BeautifulSoup(response.content, 'html.parser')
        except Exception as e:
            print(f"Error occurred while fetching {self.url}: {e}")
            self.soup = None
    
    def get_data(self):
        raise NotImplementedError("Bu metod alt sınıflar tarafından özelleştirilmelidir")