import requests
from bs4 import BeautifulSoup
import openpyxl
import re as regex

def extract_float(string):
    # Utilisation de regex non nécessaire pour extraire un float
    return float(''.join(filter(str.isdigit, string)) + '.' + ''.join(filter(str.isdigit, string.split('.')[1])))

def extract_int(string):
    # Utilisation de regex non nécessaire pour extraire un entier
    return int(''.join(filter(str.isdigit, string)))

def get_book_infos(url, book_key) :
    
    # product_page_url
    product_page_url = url

    # récupération du HTML
    r = requests.get(product_page_url)
    html = BeautifulSoup(r.content, 'html.parser')
    ws[f'A{book_key}'] = product_page_url

    # titre
    title = html.find('div', {'class':'product_main'}).find("h1").get_text()
    ws[f'C{book_key}'] = title

    # universal_product_code
    ws[f'B{book_key}'] = html.find("th", string="UPC").findNext('td').get_text()

    # Price (excl. tax)
    ws[f'D{book_key}'] = extract_float(html.find("th", string="Price (excl. tax)").findNext('td').get_text())

    # Price (incl. tax)
    ws[f'E{book_key}'] = extract_float(html.find("th", string="Price (incl. tax)").findNext('td').get_text())

    # universal_product_code
    ws[f'F{book_key}'] = extract_int(html.find("th", string="Availability").findNext('td').get_text())

    # image_url
    image_url = html.find('img').attrs['src'].replace('../..', 'https://books.toscrape.com')
    ws[f'J{book_key}'] = image_url

    # Description
    description = html.find('div', id="product_description").findNext('p').get_text()
    ws[f'G{book_key}'] = image_url

    # Category
    category = html.find('li', class_="active").findPrevious('a').get_text()
    ws[f'H{book_key}'] = category

    # rating 
    rating = html.find('p', class_="star-rating").attrs['class'][1]

    # Conversion de la notation en étoiles en entier
    rating_dict = {"Zero": 0, "One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
    ws[f'I{book_key}'] = rating_dict.get(rating)

    print(f'"{title}" téléchargé avec succès')

workbook = openpyxl.Workbook()
domain_name = "https://books.toscrape.com/"
url_website = f'{domain_name}index.html'

col_index = {
    "A" : "product_page_url",
    "B" : "universal_product_code",
    "C" : "title",
    "D" : "price_including_tax",
    "E" : "price_excluding_tax",
    "F" : "number_available",
    "G" : "product_description",
    "H" : "category",
    "I" : "review_rating",
    "J" : "image_url"
}

# récupération du HTML
website = requests.get(url_website)
html_main = BeautifulSoup(website.content, 'html.parser')

key_category = 0

for category in html_main.find('ul', class_="nav-list").find('li').find_all('li') :

    # Permet de ne capter que quelques catégories pour que ce ne soit pas trop long
    # key_category += 1
    # if key_category >= 2 :
    #     break

    category_name = category.get_text().strip()
    print('')
    print(f'Passage à la catégorie "{category_name}"')

    category_url = domain_name+(category.find('a').attrs['href'])
    ws = workbook.create_sheet(category_name)

    for key, value in col_index.items():
        ws[f'{key}1'] = value

    next_page_exists = True
    book_key = 2

    while(next_page_exists) :

        html_category = BeautifulSoup(requests.get(category_url).content, 'html.parser')

        for article in html_category.find_all('article', class_="product_pod") :
            url_article = article.find('a').attrs['href'].replace('../../..', domain_name+'catalogue')
            get_book_infos(url_article, book_key)
            book_key += 1

        next_page = html_category.find('li', class_="next")
        
        if(next_page) :
            category_url = (category_url+(next_page.find('a').attrs['href'])).replace('index.html', '')
            print('')
            print('Passage à la page suivante...')
        else :
            next_page_exists = False
            print('')

workbook.save('../data/books.xlsx')
print('Fichier enregistré avec succès.')
