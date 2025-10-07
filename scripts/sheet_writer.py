import json
import gspread
import logging
from collections import defaultdict
import os.path
from datetime import datetime
import time
from natsort import natsorted, ns

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/app.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

order_table = gspread.service_account(filename='credentials/snappy-stacker-431017-p0-08d715145a62.json')

def flatten_hierarchy(folders, result=None, level=1, stats=None, pdf_analysis_data=None, current_path=""):
    if result is None:
        result = []
    if stats is None:
        stats = {'folders': 0, 'files': defaultdict(int), 'max_level': 0}
    if pdf_analysis_data is None:
        pdf_analysis_data = {}

    for item in folders:
        if isinstance(item, dict):
            for folder_name, content in item.items():
                result.append((folder_name, level, "", "", "", "", False, True, False))  # Папка
                stats['folders'] += 1
                stats['max_level'] = max(stats['max_level'], level)
                sub_path = os.path.join(current_path, folder_name).replace('\\', '/')
                sub_folders = natsorted([sub_item for sub_item in content if isinstance(sub_item, dict)],
                                        key=lambda x: list(x.keys())[0], alg=ns.LOCALE)
                sub_files = natsorted([sub_item for sub_item in content if isinstance(sub_item, str)], alg=ns.LOCALE)
                flatten_hierarchy(sub_folders, result, level + 1, stats, pdf_analysis_data, sub_path)
                if sub_folders and sub_files:
                    result.append(("", level, "", "", "", "", True, False, False))  # Пустая строка
                for file in sub_files:
                    file_normalized = os.path.normpath(os.path.join(sub_path, file)).replace('\\', '/')
                    analysis = pdf_analysis_data.get(file_normalized, {})
                    pages = analysis.get('pages', "") if file.lower().endswith('.pdf') else ""
                    size = analysis.get('size', "")
                    orientation = analysis.get('orientation', "")
                    char_count = analysis.get('char_count', "")
                    is_archive = file.lower().endswith(('.rar', '.zip'))
                    logger.info(f"Обработка файла: {file_normalized}, страницы: {pages}, ориентация страницы: {size}, ориентация текста: {orientation}, символы: {char_count}, архив: {is_archive}")
                    result.append((file, level + 1, pages, size, orientation, char_count, False, False, is_archive))  # Файл
                    _, ext = os.path.splitext(file.lower())
                    stats['files'][ext or 'no_extension'] += 1
    return result, stats

def get_color_for_level(level):
    colors = [
        [0.9, 0.9, 0.5], [0.7, 0.9, 0.7], [0.6, 0.8, 0.9], [0.9, 0.7, 0.7],
        [0.8, 0.8, 0.8], [0.9, 0.8, 0.6], [0.7, 0.7, 0.9], [0.8, 0.9, 0.6],
        [0.9, 0.6, 0.9], [0.6, 0.9, 0.8]
    ]
    return colors[min(level, len(colors) - 1)]

def write_hierarchy_to_sheet(json_file_path, table_input):
    start_time = time.time()
    try:
        spreadsheet = order_table.open_by_url(table_input) if table_input.startswith('http') else order_table.open(
            table_input)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        worksheet_name = f"Иерархия_{timestamp}"
        worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows=1000, cols=7)

        with open(json_file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            folders = data.get('folders', [])
            files = natsorted(data.get('files', []), alg=ns.LOCALE)
            pdf_analysis_data = data.get('pdf_analysis_data', {})
            logger.info(f"PDF-анализ в json: {list(pdf_analysis_data.keys())}")
            logger.info(f"Корневые файлы: {files}")

        flat_data, stats = flatten_hierarchy(folders, pdf_analysis_data=pdf_analysis_data)

        if files:
            flat_data.append(("", 1, "", "", "", "", True, False, False))  # Пустая строка
        for file in files:
            file_normalized = os.path.normpath(file).replace('\\', '/')
            analysis = pdf_analysis_data.get(file_normalized, {})
            pages = analysis.get('pages', "") if file.lower().endswith('.pdf') else ""
            size = analysis.get('size', "")
            orientation = analysis.get('orientation', "")
            char_count = analysis.get('char_count', "")
            is_archive = file.lower().endswith(('.rar', '.zip'))
            logger.info(f"Обработка корневого файла: {file_normalized}, страницы: {pages}, ориентация страницы: {size}, ориентация текста: {orientation}, символы: {char_count}, архив: {is_archive}")
            flat_data.append((file, 1, pages, size, orientation, char_count, False, False, is_archive))
            _, ext = os.path.splitext(file.lower())
            stats['files'][ext or 'no_extension'] += 1
            stats['max_level'] = max(stats['max_level'], 1)

        logger.info(f"Обнаружено {len(flat_data)} элементов для записи.")
        logger.info(f"Статистика: {stats['folders']} папок, максимальный уровень вложенности: {stats['max_level']}")
        logger.info("Файлы по расширениям: " + ", ".join(f"{ext}: {count}" for ext, count in stats['files'].items()))

        # Подготовка данных для таблицы
        values = []
        headers = ["Имя", "Уровень", "Кол-во страниц (PDF)", "Ориентация страницы"]
        values.append(headers)

        for item, level, pages, size, orientation, char_count, is_separator, _, is_archive in flat_data:
            translation_stats = ""
            values.append(
                [f"{'  ' * level}{item}", str(level) if not is_separator else "", str(pages), size, orientation, char_count, translation_stats])

        # Записываем данные
        worksheet.update(range_name=f'A1:G{len(values)}', values=values, value_input_option="RAW")

        # Форматирование
        level_ranges = defaultdict(list)
        archive_ranges = []
        for idx, (item, level, _, _, _, _, is_separator, is_folder, is_archive) in enumerate(flat_data, start=2):
            if is_archive:
                archive_ranges.append(f'A{idx}:G{idx}')
            elif is_separator or is_folder:
                level_ranges[level].append(f'A{idx}:G{idx}')

        for level, cells in level_ranges.items():
            color = get_color_for_level(level)
            worksheet.format(cells, {
                "backgroundColor": {"red": color[0], "green": color[1], "blue": color[2]}
            })

        if archive_ranges:
            worksheet.format(archive_ranges, {
                "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 0.0}
            })

        worksheet.format('A1:G1', {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.7, "green": 0.7, "blue": 0.7}
        })

        elapsed_time = time.time() - start_time
        logger.info(f"Запись завершена в лист '{worksheet_name}'. Всего строк: {len(values)}. Время выполнения: {elapsed_time:.2f} секунд")

        summary = {
            "total_items": len(flat_data),
            "folders": stats['folders'],
            "max_level": stats['max_level'],
            "files_by_extension": dict(stats['files']),
            "pdf_analysis_data": pdf_analysis_data
        }
        return f"Иерархия успешно записана в лист '{worksheet_name}'.", summary

    except gspread.SpreadsheetNotFound:
        elapsed_time = time.time() - start_time
        logger.error(f"Таблица не найдена: {table_input}. Время выполнения: {elapsed_time:.2f} секунд")
        raise Exception(f"Ошибка: Таблица '{table_input}' не найдена")
    except Exception as e:
        elapsed_time = time.time() - start_time
        logger.error(f"Ошибка: {e}. Время выполнения: {elapsed_time:.2f} секунд")
        raise