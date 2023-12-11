from jinja2 import Environment, FileSystemLoader
from input_data import data as input_data
import requests
import pandas as pd

def get_photo_url(vendor_code: str) -> str:
    if vendor_code.startswith('HB'):
        return f'http://pravim.by/upload_price/Haiba/photo/{vendor_code}_photo.jpg'
    elif vendor_code.startswith('CN'):
        return f'http://pravim.by/upload_price/Cron/photo/{vendor_code}_photo.jpg'
    return None

def get_pricelist(urls=[
    'https://docs.google.com/spreadsheets/d/1MxIv111Z1NnP6qbcJ_QIL5FM8IWlp4qAn3ifCUGZooM/export?format=csv&id=1MxIv111Z1NnP6qbcJ_QIL5FM8IWlp4qAn3ifCUGZooM&gid=1949199208',
    'https://docs.google.com/spreadsheets/d/1MxIv111Z1NnP6qbcJ_QIL5FM8IWlp4qAn3ifCUGZooM/export?format=csv&id=1MxIv111Z1NnP6qbcJ_QIL5FM8IWlp4qAn3ifCUGZooM&gid=125708755'
                ]):
    df = pd.DataFrame()
    for url in urls:
        response = requests.get(url)
        with open(f'price_{urls.index(url)}.csv', 'wb') as file:
            file.write(response.content)
        df = pd.concat([df, pd.read_csv(url, encoding='utf-8', usecols=['Артикул','РРЦ с НДС, BYN (справочно)', 'Фото'], index_col='Артикул')])
    return df

def main():  
    pricelist = get_pricelist()  
    env = Environment(loader=FileSystemLoader("templates/"))
    env.filters['get_photo_url'] = get_photo_url
    template= env.from_string(open('templates/matrix.html').read())

    for row in input_data:
        for i, item in enumerate(row):
            values = pricelist.loc[item]
            row[i] = {
                'vendor_code': item,
                'photo': values.get('Фото'),
                'price': values.get('РРЦ с НДС, BYN (справочно)'), 
            }

    with open('index.html', 'w', encoding='utf-8') as file:
        file.write(template.render(input_data = input_data))
        
if __name__ == "__main__":
    main()