from html2excel import ExcelParser



if __name__ == '__main__':
    input_file = 'index.html'
    output_file = 'index.xlsx'

    parser = ExcelParser(input_file)
    parser.to_excel(output_file)