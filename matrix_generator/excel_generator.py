
import requests
import pandas as pd
import os
from PIL import Image as PIL_Image

from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import  XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU as p2e, pixels_to_points
from openpyxl.styles import Border, Side, Alignment, Font

from io import BytesIO

IMG_OFFSET = 2 #Отступ в пикселях от угла ячейки к которой приязано изображение (использовать меньше 2 не рекомендуется)

ATRR_NAMES = ['Торговая марка', 'Артикул', 'Материал', 'Цвет', 'Розничная цена (ориентировочно), BYN']


def resized_image(size,img_url):
    cached_image_path = f'cached_images/{size}/{img_url.split("/")[-1]}'
    if not os.path.exists(cached_image_path):
        os.makedirs(os.path.join(*cached_image_path.split('/')[:-1]),exist_ok=True)
        
        image = PIL_Image.open(requests.get(img_url, stream = True).raw)
        image.thumbnail((size,size), PIL_Image.LANCZOS)
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

def parse_matrix (file_path = 'matrix_generator/matrix.txt'):
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


matrix, customer, img_size = parse_matrix()\
    
def generate_excel_file(matrix, customer, img_size):
    pricelist = get_pricelist()  
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

    wb = Workbook()
    ws = wb.active
    ws.sheet_view.showGridLines = False
    bd = Side(style='thin', color="000000")
    border_style = Border(left=bd, top=bd, right=bd, bottom=bd)
    alignment=Alignment(horizontal='center')

    current_row = 2
    for matrix_row in matrix:
        current_collumn = 3
        #Ширина ячейки с наименованиями атрибутов 
        ws.column_dimensions[get_column_letter(2)].width = 38

        #Цикл по "строкам матрицы"
        for i, item in enumerate(matrix_row):
            # Вставка наименований атрибутов
            for name_i, name in enumerate(ATRR_NAMES):
                ws.cell(row = current_row+name_i+1, column = 2, value= name).border = border_style
                
            
            # Цикл по атрибутам    
            for j, prop in enumerate(list(item.values())):
                if j == 0:
                    img = Image(prop)
                    
                    #Позиционирование изображение + добавление отступа для того чтобы были видны границы разметки
                    h, w = img.height, img.width
                    size = XDRPositiveSize2D(p2e(w), p2e(h))
                    column = current_collumn
                    row = current_row+j
                    marker = AnchorMarker(col=column-1, colOff=p2e(IMG_OFFSET), row=row-1, rowOff=p2e(IMG_OFFSET))
                    img.anchor = OneCellAnchor(_from=marker, ext=size)

                    ws.add_image(img)
                    
                    cell = ws.cell(row = current_row+j, column = current_collumn)
                    cell.border = border_style
                    cell.alignment = alignment
                    
                    #Задаем ширину и высоту строки с изображением (числа подобраны исходя из размеров)
                    if not ws.column_dimensions[get_column_letter(current_collumn)].width or ws.column_dimensions[get_column_letter(current_collumn)].width < (w+2*IMG_OFFSET)/7:
                        ws.column_dimensions[get_column_letter(current_collumn)].width = (w+2*IMG_OFFSET)/7
                    if not ws.row_dimensions[current_row+j].height or ws.row_dimensions[current_row+j].height < pixels_to_points(h+2*IMG_OFFSET+2):                      
                        ws.row_dimensions[current_row+j].height = pixels_to_points(h+2*IMG_OFFSET+2)
                else:   
                    cell = ws.cell(row = current_row+j, column = current_collumn, value = prop)
                    cell.alignment = alignment
                    cell.border = border_style
                    
            current_collumn +=1
        current_row = current_row + len(item) +1
        
    # Save the workbook
    # wb.save('product_table_with_images.xlsx')
    virtual_workbook = BytesIO()
    wb.save(virtual_workbook)
    return virtual_workbook.getbuffer()

# if __name__ == "__main__":
#     main()