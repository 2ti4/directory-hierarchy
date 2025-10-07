import logging
import platform
import locale
from flask import Flask, request, render_template, jsonify
from pathlib import Path
import json
import os.path
from scripts.archive_struct import get_directory_structure
from scripts.sheet_writer import write_hierarchy_to_sheet
from scripts.analysis_utils import count_pdf_pages, get_page_size
from concurrent.futures import ThreadPoolExecutor, as_completed
from natsort import natsorted, ns
from datetime import datetime

# Устанавливаем локаль для корректной сортировки кириллицы
try:
    locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')
except locale.Error:
    logging.warning("Локаль ru_RU.UTF-8 не поддерживается, используется стандартная локаль")

# Настройка логирования
LOG_DIR = Path('logs')
LOG_FILE = LOG_DIR / 'app.log'
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
OUTPUT_FOLDER = Path('output')
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
OUTPUT_FOLDER.mkdir(exist_ok=True)

def convert_path_to_current_platform(windows_path):
    """
    Преобразует путь в формат, подходящий для текущей платформы (Windows или WSL).
    """
    windows_path = windows_path.strip()
    if not windows_path:
        return None
    is_windows = platform.system() == "Windows"
    is_wsl = platform.system() == "Linux" and "microsoft" in platform.uname().release.lower()
    if is_wsl:
        if len(windows_path) > 1 and windows_path[0].isalpha() and windows_path[1] == ':':
            drive_letter = windows_path[0].lower()
            rest_of_path = windows_path[2:].lstrip('\\').replace('\\', '/')
            wsl_path = f"/mnt/{drive_letter}/{rest_of_path}"
            return wsl_path
        return windows_path
    elif is_windows:
        if windows_path.startswith('/mnt/') and len(windows_path) > 6 and windows_path[5].isalpha():
            drive_letter = windows_path[5].upper()
            rest_of_path = windows_path[6:].lstrip('/').replace('/', '\\')
            windows_path = f"{drive_letter}:\\{rest_of_path}"
            return windows_path
        return windows_path.replace('/', '\\')
    return windows_path

@app.route('/', methods=['GET', 'POST'])
def process_directory():
    if request.method == 'POST':
        server_path = request.form.get('server_path', '').strip()
        table_input = request.form.get('table_input', '').strip()

        if not server_path:
            return jsonify({"error": "Не указан путь к папке"}), 400

        server_path = convert_path_to_current_platform(server_path)
        server_path_obj = Path(server_path)
        if not server_path_obj.exists() or not server_path_obj.is_dir():
            return jsonify({"error": f"Указанная папка не существует или не является директорией: {server_path}"}), 400

        try:
            # Получаем структуру папки
            structure = get_directory_structure(server_path)
            if structure and isinstance(structure, list) and "error" in structure[0]:
                return jsonify({"error": structure[0]["error"]}), 400

            # Разделяем папки и файлы
            folders = natsorted([item for item in structure if isinstance(item, dict)], key=lambda x: list(x.keys())[0], alg=ns.LOCALE)
            files = natsorted([item for item in structure if isinstance(item, str)], alg=ns.LOCALE)

            logger.info(f"Отсортированные файлы: {files}")

            # Собираем список файлов для анализа (PDF, DOCX, XLSX)
            files_to_analyze = []
            for root, dirs, files_walk in server_path_obj.walk():
                dirs[:] = natsorted(dirs, alg=ns.LOCALE)
                files_walk[:] = natsorted(files_walk, alg=ns.LOCALE)
                for file in files_walk:
                    if file.lower().endswith(('.pdf', '.docx', '.xlsx')):
                        files_to_analyze.append(root / file)

            pdf_analysis_data = {}
            if files_to_analyze:
                max_workers = min(8, len(files_to_analyze))
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    future_to_file = {executor.submit(count_pdf_pages, file_path): file_path for file_path in files_to_analyze}
                    for future in as_completed(future_to_file):
                        file_path, pages = future.result()
                        relative_path = os.path.normpath(str(file_path.relative_to(server_path_obj))).replace('\\', '/')
                        if relative_path.startswith('6. НВК/'):
                            relative_path = relative_path[len('6. НВК/'):]
                        size_str = get_page_size(file_path) if file_path.suffix.lower() == '.pdf' else ""
                        pdf_analysis_data[relative_path] = {
                            "pages": pages,
                            "size": size_str
                        }
                        logger.info(f"Добавлен файл: {relative_path} с {pages} страницами")

            # Добавляем файлы без анализа
            for root, dirs, files_walk in server_path_obj.walk():
                for file in files_walk:
                    if not file.lower().endswith(('.pdf', '.docx', '.xlsx')):
                        relative_path = os.path.normpath(str((root / file).relative_to(server_path_obj))).replace('\\', '/')
                        if relative_path.startswith('6. НВК/'):
                            relative_path = relative_path[len('6. НВК/'):]
                        pdf_analysis_data[relative_path] = {
                            "pages": "",
                            "size": ""
                        }
                        logger.info(f"Добавлен файл без анализа: {relative_path}")

            pdf_analysis_data = dict(natsorted(pdf_analysis_data.items(), key=lambda x: x[0], alg=ns.LOCALE))
            logger.info(f"Итоговый pdf_analysis_data: {list(pdf_analysis_data.keys())}")

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            json_path = app.config['OUTPUT_FOLDER'] / f'inventory_structure_{timestamp}.json'
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump({
                    "root_folder": server_path_obj.name,
                    "folders": folders,
                    "files": files,
                    "pdf_analysis_data": pdf_analysis_data
                }, f, indent=4, ensure_ascii=False)

            hierarchy_result = {"message": "Иерархия обработана", "json_path": str(json_path)}
            if table_input:
                try:
                    message, summary = write_hierarchy_to_sheet(json_path, table_input)
                    json_path.unlink(missing_ok=True)
                    return jsonify({"hierarchy": hierarchy_result, "message": message, "summary": summary})
                except Exception as e:
                    json_path.unlink(missing_ok=True)
                    logger.error(f"Ошибка записи в Google Sheets: {str(e)}")
                    return jsonify({"hierarchy": hierarchy_result, "error": str(e)}), 400

            json_path.unlink(missing_ok=True)
            return jsonify({"hierarchy": hierarchy_result})

        except Exception as e:
            logger.error(f"Ошибка обработки папки {server_path}: {e}")
            return jsonify({"error": str(e)}), 400

    return render_template('input_path.html')

if __name__ == '__main__':
    logger.info(f"Запуск приложения на платформе: {platform.system()}")
    try:
        app.run(debug=True, host='0.0.0.0', port=5002)
    except Exception as e:
        logger.error(f"Ошибка запуска сервера: {str(e)}")
        raise