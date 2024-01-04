from jinja2 import Environment, FileSystemLoader
import requests
import pandas as pd
from PIL import Image
import os

def resized_image(size,img_url):
    cached_image_path = f'cached_images/{size}/{img_url.split("/")[-1]}'
    if not os.path.exists(cached_image_path):
        os.makedirs(os.path.join(*cached_image_path.split('/')[:-1]),exist_ok=True)
        
        image = Image.open(requests.get(img_url, stream = True).raw)
        image.thumbnail((size,size), Image.LANCZOS)
        image.save(cached_image_path)
        print(cached_image_path)
    return cached_image_path


def get_pricelist(urls=[
    'https://docs.google.com/spreadsheets/d/1MxIv111Z1NnP6qbcJ_QIL5FM8IWlp4qAn3ifCUGZooM/export?format=csv&id=1MxIv111Z1NnP6qbcJ_QIL5FM8IWlp4qAn3ifCUGZooM&gid=1949199208',
    'https://docs.google.com/spreadsheets/d/1MxIv111Z1NnP6qbcJ_QIL5FM8IWlp4qAn3ifCUGZooM/export?format=csv&id=1MxIv111Z1NnP6qbcJ_QIL5FM8IWlp4qAn3ifCUGZooM&gid=125708755'
                ]):
    df = pd.DataFrame()
    for url in urls:
        response = requests.get(url)
        with open(f'price_{urls.index(url)}.csv', 'wb') as file:
            file.write(response.content)
        df = pd.concat([df, pd.read_csv(url, encoding='utf-8', usecols=['Артикул','РРЦ с НДС, BYN (справочно)', 'Фото', 'Бренд', 'Материал', 'Цвет'], index_col='Артикул')])
    return df

def parse_matrix (file_path = 'matrix.txt'):
    matrix = []
    img_size = None
    customer = None
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file.readlines():
            if line.startswith('Контрагент'): 
                customer = line.split(':')[-1].strip()
            elif line.startswith('Размер изображения'): 
                try:
                    img_size = int(line.split(':')[-1].strip())
                except ValueError:
                    img_size = None
            else:
                matrix.append ([x.strip() for x in line.strip().split(',')])
                
                
    return matrix, customer, img_size

def get_template_name():
    template_query = '_________________\nВыберите template\n'
    templates_names = []
    for index, file in enumerate(os.listdir('./templates')):
        template_query += f"{index+1} - {file}\n"
        templates_names.append(file)
    template_query += 'Ваш выбор: '
    
    template_selection = None
    while not template_selection:
        try: 
            template_selection = int(input(template_query))
            if template_selection > len(templates_names) or template_selection < 1:
                template_selection = None
                raise ValueError ('Введено неверное значение.')   
        except ValueError as ex:
            print(ex)
            print('Попробуйте ещё раз')     

    return templates_names[template_selection-1]
def main():  
    matrix, customer, img_size = parse_matrix()
    print(matrix)
    template_name = get_template_name()
    print (f'Выбрано: {template_name}')
    pricelist = get_pricelist()  
    env = Environment(loader=FileSystemLoader("templates/"))
    # template= env.from_string(open('templates/matrix.html', encoding='utf-8').read()) #матрица на grid
    template= env.from_string(open(f'templates/{template_name}', encoding='utf-8').read()) #матрица на table
    

    for row in matrix:
        for i, item in enumerate(row):
            values = pricelist.loc[item]
            row[i] = {
                'photo': resized_image(img_url=values.get('Фото'), size=img_size) if img_size else values.get('Фото'),
                'brand': values.get('Бренд'),
                'vendor_code': item,
                'material': values.get('Материал'),
                'color': values.get('Цвет'),
                'price': values.get('РРЦ с НДС, BYN (справочно)'),
           
            }

    with open('index.html', 'w', encoding='utf-8') as file:
        file.write(template.render(input_data = matrix, customer = customer, ))
    print ('Результат сохранен в index.html')
if __name__ == "__main__":
    main()