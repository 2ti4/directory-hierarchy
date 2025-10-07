import logging
import pdfplumber

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

def count_pdf_pages(file_path):
    """Подсчёт страниц в PDF-файле."""
    try:
        logger.info(f"Попытка открыть PDF: {file_path}")
        with pdfplumber.open(file_path) as pdf:
            page_count = len(pdf.pages)
            logger.info(f"Успешно подсчитано {page_count} страниц для файла: {file_path}")
            return file_path, page_count
    except Exception as e:
        logger.error(f"Ошибка при обработке {file_path}: {str(e)}")
        return file_path, 0

def get_page_size(file_path):
    """Определяет ориентацию страницы: альбомный или книжный."""
    try:
        with pdfplumber.open(file_path) as pdf:
            page = pdf.pages[0]
            width, height = page.width, page.height
            if width is None or height is None:
                return "Error: Unable to get size"
            return "альбомный" if width > height else "книжный"
    except Exception as e:
        logger.error(f"Ошибка получения ориентации страницы для {file_path}: {str(e)}")
        return f"Error: {str(e)}"