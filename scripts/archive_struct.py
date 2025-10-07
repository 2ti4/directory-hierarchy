from pathlib import Path
from collections import deque
from natsort import natsorted, ns
import rarfile
import zipfile

# Функция get_archive_structure больше не используется, но оставлена закомментированной
"""
def get_archive_structure(archive_path):
    \"\"\"Получает структуру архива (RAR/ZIP) с натуральной сортировкой, игнорируя файлы .db.\"\"\"
    structure = []
    archive_type = archive_path.suffix.lower()
    try:
        opener = rarfile.RarFile if archive_type == '.rar' else zipfile.ZipFile
        with opener(archive_path) as archive:
            temp_dict = {}
            for item in archive.infolist():
                if hasattr(item, 'is_dir') and item.is_dir():
                    continue
                if item.filename.lower().endswith('.db'):
                    continue
                path_parts = item.filename.split('/')
                current = temp_dict
                for part in path_parts[:-1]:
                    if part and part not in current:
                        current[part] = {}
                    if part:
                        current = current[part]
                if path_parts[-1]:
                    current[path_parts[-1]] = "file"

            def dict_to_list(d):
                result = []
                folders = [(k, v) for k, v in d.items() if isinstance(v, dict) and v]
                files = [k for k, v in d.items() if not isinstance(v, dict)]
                folders = natsorted(folders, key=lambda x: x[0], alg=ns.LOCALE)
                files = natsorted(files, alg=ns.LOCALE)
                for key, value in folders:
                    result.append({key: dict_to_list(value)})
                result.extend(files)
                return result

            structure = dict_to_list(temp_dict)
        return structure
    except Exception as e:
        return [{"error": str(e)}]
"""

def get_directory_structure(path):
    """Итеративно получает структуру директории с натуральной сортировкой, игнорируя файлы .db."""
    path = Path(path)
    structure = []
    stack = deque([(path, structure)])

    try:
        while stack:
            current_path, current_structure = stack.pop()
            items = natsorted(current_path.iterdir(), key=lambda x: x.name, alg=ns.LOCALE)
            for item in items:
                if item.is_dir():
                    sub_structure = []
                    current_structure.append({item.name: sub_structure})
                    stack.append((item, sub_structure))
                elif item.suffix.lower() != '.db':
                    current_structure.append(item.name)
    except PermissionError:
        structure.append({"error": "Permission denied"})
    except Exception as e:
        structure.append({"error": str(e)})

    return structure