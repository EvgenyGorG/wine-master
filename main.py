import datetime
from pathlib import Path
from collections import defaultdict

from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas as pd

from jinja2 import Environment, FileSystemLoader, select_autoescape


def get_correct_form(period_of_service):
    period_of_service = str(period_of_service)
    if (
            len(period_of_service) >= 2 and
            period_of_service[-2:] in ['11', '12', '13', '14']
    ):
        return f'{period_of_service} лет'
    elif period_of_service[-1] == '1':
        return f'{period_of_service} год'
    elif period_of_service[-1] in ['2', '3', '4']:
        return f'{period_of_service} года'
    else:
        return f'{period_of_service} лет'


def get_wine_catalog():
    wine_catalog_folder = Path('wine_catalog')

    if not wine_catalog_folder.exists():
        raise FileNotFoundError(f"Папка {wine_catalog_folder} не найдена")

    wine_catalogs = list(wine_catalog_folder.glob('*.xlsx'))

    if not wine_catalogs:
        raise FileNotFoundError(f"Файлы .xlsx не найдены в папке {wine_catalog_folder}")

    if len(wine_catalogs) > 1:
        raise ValueError(f'В папке {wine_catalog_folder}'
                        f' найдено больше одного файла,'
                        f' ожидается один.')

    return wine_catalogs[0]


def format_wines_info(wines_info):
    formatted_wines_info = defaultdict(list)

    for wine_info in wines_info:
        category = wine_info.pop('Категория')
        formatted_wines_info[category].append(wine_info)

    return formatted_wines_info


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')

    current_year = datetime.datetime.now().year
    year_of_foundation = 1920
    period_of_service = current_year - year_of_foundation

    wine_catalog = get_wine_catalog()

    excel_data_df = pd.read_excel(wine_catalog, na_values=' ', keep_default_na=False)
    wines_info = excel_data_df.to_dict(orient='records')

    formatted_wines_info = format_wines_info(wines_info)

    rendered_page = template.render(
        formatted_wines_info=formatted_wines_info,
        period_of_service=get_correct_form(period_of_service)
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()