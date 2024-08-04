from bs4 import BeautifulSoup as BS
import requests
from openpyxl import Workbook


def get_html(url):
    headers = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36'}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.text
    return None

def get_glide_link(html):
    soup = BS(html, 'html.parser')
    content = soup.find('div', class_='wrapper')
    main = content.find('main', class_='content')
    text = main.find('article', class_='text')
    post = text.find_all('section', class_='post-list')
    links = []
    for p in post:
        img = p.find('div', class_='img-link')
        link = img.find('a').get('href')
        links.append(link)
    return links


def get_data(html):
    soup = BS(html, 'html.parser')
    content = soup.find('article', class_='text')
    post = content.find('section', class_='post-singl')
    name = post.find('h1').text.strip()
    print(name)
    table = post.find('table', class_='table-tag')
    country = table.find('tbody').text.strip()
    print(country)
    tbody = post.find('tbody', class_='tbody-sin')
    time = tbody.find('tr').text.strip()
    print(time)
    canal = tbody.find('tr', class_='linkpost')
    canal = canal.text.strip() if canal else 'no canal'
    perevod = tbody.find('tr', class_='perevod').text.strip()
    print(perevod)
    person = tbody.find('tr', class_='person')
    person = person.text.strip() if person else 'no person'
    description = content.find('div', class_='description').find('p').text.strip()
    print(description)
    raiting = content.find('div', class_='rt-bl')
    rait = raiting.find('div', class_='unit-rating').find('span').text.strip()
    print(rait)   
    commlist = content.find('div', class_='comment-full')   
    comment = commlist.find('ul', class_='commentlist')
    comm = comment.find('div', class_='ct-text clearfix')
    comm = comm.find('p').text.strip() if comm else 'no comm'



    data = {
        'name': name,
        'country and zhanr': country,
        'time': time,
        'canal': canal,
        'perevod': perevod,
        'person': person,
        'description': description,
        'raiting': rait,
        'comment': comm
        
    }
    
    return data

def last_page(html):
    soup = BS(html, 'html.parser')
    page = soup.find('div', class_='art-pager')
    pages = page.find_all('a', class_='page-numbers')
    last_page = pages[-2].text
    return int(last_page)
    

def save_to_excel(data):
    workbook = Workbook()
    sheet = workbook.active
    sheet['A1'] = 'Название'
    sheet['B1'] = 'Страна и жанры'
    sheet['C1'] = 'Продолжительность'
    sheet['D1'] = 'Канал'
    sheet['E1'] = 'Перевод'
    sheet['F1'] = 'В ролях'
    sheet['G1'] = 'Описание'
    sheet['H1'] = 'Рейтинг'
    sheet['I1'] = 'Комментарий'


    for i,item in enumerate(data,2):
        sheet[f'A{i}'] = item['name']
        sheet[f'B{i}'] = item['country and zhanr']
        sheet[f'C{i}'] = item['time']
        sheet[f'D{i}'] = item['canal']
        sheet[f'E{i}'] = item['perevod']
        sheet[f'F{i}'] = item['person']
        sheet[f'G{i}'] = item['description']
        sheet[f'H{i}'] = item['raiting']
        sheet[f'I{i}'] = item['comment']
    
    workbook.save('dorama_data.xlsx')   


def main():
    URL = 'https://doramy.club/'
    html = get_html(url=URL)
    page = last_page(html)
    data = []
    for i in range(1, 4):
        page_url = URL + f'page/{i}'
        page_html = get_html(page_url)
        links = get_glide_link(page_html)
        for link in links:
            posts_links = get_html(url=link)
            data.append(get_data(html=posts_links))

    save_to_excel(data)
            
       


if __name__ == '__main__':
    main()