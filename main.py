from http.server import HTTPServer, SimpleHTTPRequestHandler
import datetime
from jinja2 import Environment, FileSystemLoader, select_autoescape
import pandas
import collections


excel_data_df = pandas.read_excel(r'wine3.xlsx',
                                  sheet_name='katalog',
                                  usecols=['Категория',
                                           'Название',
                                           'Сорт',
                                           'Цена',
                                           'Картинка',
                                           'Акция'],
                                  na_values=True,
                                  keep_default_na='')
payload = {val1: val2.to_dict(orient='records') for val1, val2 in excel_data_df.groupby("Категория")}
katalog = collections.Counter(payload)


def year_format(how_old):
    if how_old > 100:
        if (how_old % 100) in (11, 12, 13, 14):
            return f'{how_old} лет'
        elif (how_old % 10) in (5, 6, 7, 8, 9, 0):
            return f'{how_old} лет'
        elif (how_old % 10) == 1:
            return f'{how_old} год'
        elif (how_old % 10) in (2, 3, 4):
            return f'{how_old} года'
    elif how_old in (10, 11, 12, 13, 14):
        return f'{how_old} лет'
    elif (how_old % 10) in (5, 6, 7, 8, 9, 0):
        return f'{how_old} лет'
    elif (how_old % 10) == 1:
        return f'{how_old} год'
    else:
        return f'{how_old} года'


env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')

how_old = datetime.datetime.now().year - 1920

rendered_page = template.render(
    how_old_text=year_format(how_old),
    katalog=katalog
)


with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
