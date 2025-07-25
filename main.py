import argparse
import datetime
from pathlib import Path
from collections import defaultdict

from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas as pd

from environs import Env
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


def wine_catalog_check(wine_catalog):
    if not Path(wine_catalog).exists():
        raise FileNotFoundError("Файл по данному пути не существует.\n"
                                "Убедитесь, что указали правильный путь"
                                " к файлу и попробуйте еще раз.")

    if not Path(wine_catalog).suffix.lower() == '.xlsx':
        raise FileNotFoundError("Файл по данному пути существует,"
                                " но его расширение не соответствует"
                                " .xlsx")

    else:
        return wine_catalog


def format_wines_info(wines_info):
    formatted_wines_info = defaultdict(list)

    for wine_info in wines_info:
        category = wine_info.pop('Категория')
        formatted_wines_info[category].append(wine_info)

    return formatted_wines_info


def main():
    env = Env()
    env.read_env()

    wine_catalog = env.str(
        'WINE_CATALOG_PATH', default='wine_catalog/wine_catalog.xlsx'
    )

    parser = argparse.ArgumentParser(
        description='Отрисовка сайта "Новое русское вино"')
    parser.add_argument(
        '-p',
        '--catalog_path',
        default=wine_catalog,
        type=Path,
        help='Относительный путь к файлу с каталогом напитков'
    )
    args = parser.parse_args()
    wine_catalog = wine_catalog_check(args.catalog_path)

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')

    current_year = datetime.datetime.now().year
    year_of_foundation = 1920
    period_of_service = current_year - year_of_foundation

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