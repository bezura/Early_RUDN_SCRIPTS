import json
import re

import pandas as pd

YEAR_OF_DOCS = 2023  # Год который указан в портфолио направления
EXPORT_RESULTS_FILE = "export.txt"
DIRECTIONS_ALIAS = {  # Расшифровка ОУП
    "ИРЯ": "Институт русского языка",
    "ВШУ": "Высшая школа управления",
    "ФГСН": "Факультет гуманитарных и социальных наук",
    "ФФМиЕН": "Факультет физико-математических и естественных наук",
    "ИЭ": "Институт экологии",
    "АТИ": "Аграрно-технологический университет",
    "ФФ": "Филологический факультет",
    "ИБХТН": "Институт биохимической технологии и нанотехнологии",
    "ИМЭБ": "Институт мировой экономики и бизнеса",
    "ЭФ": "Экономический факультет",
    "ЮИ": "Юридический институт",
    "ИИЯ": "Институт иностранных языков",
}
EXCEL_FILE_NAME = "data.xlsx"
EXCEL_COLUMN_NAMES = {
    'code': 'ОКСО',
    'direction_name': 'Направления',
    'faculty': 'ОУП',
    'place_count': 'Количество бесплатных мест'
}


def format_place_count(count):
    last_digit = int(count) % 100
    if last_digit == 0:
        return f"{count} бесплатных мест"
    if last_digit == 1:
        return f"{count} бесплатное место"
    elif 2 <= last_digit <= 4:
        return f"{count} бесплатных места"
    elif 5 <= last_digit <= 19:
        return f"{count} бесплатных мест"

    raise f"Чёт не склоняется... Да-да, реально, вот что это вообще - {count}?"


result_object = {"faculties": []}

df = pd.read_excel(EXCEL_FILE_NAME, dtype=str)

grouped = df.groupby(EXCEL_COLUMN_NAMES['faculty'])

for faculty, data in grouped:
    if not DIRECTIONS_ALIAS[faculty]:
        raise f"В DIRECTIONS_ALIAS нет расшифровки для ОУП-а {faculty}"

    data_faculty = {
        "name": DIRECTIONS_ALIAS[faculty],
        "url_results": f"/assets/docs/result/{faculty}.pdf",
        "directions": []
    }

    for _, row in data.iterrows():
        direction_name = (row[EXCEL_COLUMN_NAMES['direction_name']])
        direction = dict(code=(row[EXCEL_COLUMN_NAMES['code']]),
                         name=row[EXCEL_COLUMN_NAMES['direction_name']],
                         place_count=format_place_count(row[EXCEL_COLUMN_NAMES['place_count']]),
                         format="Очная",
                         portfolio_url=f"/assets/docs/{faculty}/{(row[EXCEL_COLUMN_NAMES['code']])}"
                                       f"_{(direction_name).upper().replace(' ', '_')}"
                                       f"_{YEAR_OF_DOCS}_{faculty}.pdf")
        data_faculty["directions"].append(direction)
    result_object["faculties"].append(data_faculty)

with open(EXPORT_RESULTS_FILE, "w", encoding="utf-8") as f:
    f.write(re.sub(r'"(.*?)"(?=:)', r'\1', json.dumps(result_object, ensure_ascii=False, indent=4)).replace("\"", "\'"))

print(f"Файл сохранён как {EXPORT_RESULTS_FILE}")
