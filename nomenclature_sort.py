#!/usr/bin/env python3
"""
Автоматическая сортировка архивов номенклатуры по категориям 1С УНФ.

Логика:
  1. Читает Структура.xlsx (выгрузка из 1С УНФ), строит соответствие
     код номенклатуры -> путь категории, используя ГРУППИРОВКУ (Outline)
     Excel-листа (тот самый +/- слева от номеров строк).
  2. Для каждого файла в SOURCE извлекает код из имени (формат НФ-XXXXXXXX).
  3. Находит категорию по коду, перемещает файл в DEST/<категория>/,
     создавая подпапки при необходимости.
  4. Если в целевой папке уже есть файл с таким именем — заменяет его.
  5. Файлы с неверным форматом имени или без совпадения в прайс-листе
     НЕ ТРОГАЕТ, только логирует в unprocessed_report.csv.

Использование:
  python nomenclature_sort.py                 # обычный запуск, реально переносит файлы
  python nomenclature_sort.py --dry-run        # только показать, что было бы сделано
  python nomenclature_sort.py --dump-map       # вывести карту код->категория и выйти
                                                  (для проверки, что прайс-лист читается верно)

Вызывается из n8n через узел Execute Command по расписанию (9:00, 13:00, 16:00).
"""

import re
import csv
import shutil
import sys
import json
import argparse
from pathlib import Path
from datetime import datetime

try:
    import openpyxl
except ImportError:
    print(json.dumps({"status": "error", "message": "Не установлен openpyxl. Установите: pip install openpyxl"}, ensure_ascii=False))
    sys.exit(1)

# --- Конфигурация путей ---
# ТЕСТОВЫЙ РЕЖИМ: локальная песочница вместо боевых сетевых папок.
# Когда логика проверена и всё работает верно — раскомментируйте боевые
# пути ниже и закомментируйте тестовые.
WORKSPACE = Path(r"C:\Users\User\Documents\Hermes agent files\Сортировка_Workspace")
SOURCE = WORKSPACE / "Nomenklaturshchik"
DEST = WORKSPACE / "GotovayaNomenklatura"
PRICE_LIST = WORKSPACE / "PriceList" / "Структура.xlsx"
LOG_DIR = Path(r"C:\Users\User\Documents\Hermes agent files")

# --- Боевые пути (раскомментировать, когда тест пройден) ---
# SOURCE = Path(r"\\KometaNAS\Производство\Номенклатура\Номенклатурщик")
# DEST = Path(r"\\KometaNAS\Производство\Номенклатура\Готовая номенклатура")
# PRICE_LIST = Path(r"C:\Users\User\Documents\Hermes agent files\Структура.xlsx")
TRANSFER_LOG = LOG_DIR / "transfer_log.csv"
ISSUES_LOG = LOG_DIR / "unprocessed_report.csv"

# Единственный допустимый формат имени. Всё остальное (например, УУНФ-)
# считается ошибкой и не обрабатывается.
CODE_PATTERN = re.compile(r"^(НФ-\d{8})")

# Столбец с кодом номенклатуры (индекс с 1). ПРОВЕРЬТЕ по вашему файлу
# через --dump-map перед первым реальным запуском.
CODE_COLUMN = 2
# Столбец, из которого брать текст заголовка категории, если строка —
# заголовок, а не код (обычно название лежит в том же или соседнем столбце).
NAME_COLUMN = 5


def build_category_map(xlsx_path: Path) -> dict[str, str]:
    """Строит {код: относительный_путь_категории} по группировке Excel."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active

    mapping: dict[str, str] = {}
    stack: dict[int, str] = {}  # уровень вложенности -> название категории

    for row in ws.iter_rows(min_row=1):
        code_cell = row[CODE_COLUMN - 1]
        level = ws.row_dimensions[code_cell.row].outline_level
        value = str(code_cell.value) if code_cell.value else ""
        match = CODE_PATTERN.match(value)

        if match:
            # строка с кодом номенклатуры — собрать путь из всех родительских
            # заголовков (уровни строго меньше текущего)
            path_parts = [stack[lvl] for lvl in sorted(stack) if lvl < level]
            mapping[match.group(1)] = "/".join(path_parts) if path_parts else "Без категории"
        else:
            # заголовок категории — обновить стек, сбросить более глубокие уровни
            name_cell = row[NAME_COLUMN - 1].value or code_cell.value
            if name_cell and str(name_cell).strip():
                stack[level] = str(name_cell).strip()
                for lvl in [l for l in stack if l > level]:
                    del stack[lvl]

    return mapping


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="ничего не перемещать, только показать план")
    parser.add_argument("--dump-map", action="store_true", help="вывести карту код->категория и выйти")
    args = parser.parse_args()

    if not PRICE_LIST.exists():
        print(json.dumps({"status": "error", "message": f"Прайс-лист недоступен: {PRICE_LIST}"}, ensure_ascii=False))
        sys.exit(1)

    category_map = build_category_map(PRICE_LIST)
    if not category_map:
        print(json.dumps({"status": "error", "message": "Прайс-лист прочитан, но карта категорий пуста — проверьте CODE_COLUMN/NAME_COLUMN"}, ensure_ascii=False))
        sys.exit(1)

    if args.dump_map:
        for code, path in sorted(category_map.items()):
            print(f"{code}\t{path}")
        print(f"\nВсего кодов в прайс-листе: {len(category_map)}", file=sys.stderr)
        return

    if not SOURCE.exists():
        print(json.dumps({"status": "error", "message": f"Папка-источник недоступна: {SOURCE}"}, ensure_ascii=False))
        sys.exit(1)

    transferred = []
    issues = []

    for file in sorted(SOURCE.iterdir()):
        if not file.is_file():
            continue

        match = CODE_PATTERN.match(file.name)
        if not match:
            issues.append((file.name, "неверный формат имени (ожидался НФ-XXXXXXXX)"))
            continue

        code = match.group(1)
        category_path = category_map.get(code)

        if category_path is None:
            issues.append((file.name, f"код {code} не найден в прайс-листе"))
            continue

        target_dir = DEST / category_path
        target_file = target_dir / file.name
        replaced = target_file.exists()

        if args.dry_run:
            transferred.append((file.name, str(target_file.relative_to(DEST)), "ЗАМЕНА" if replaced else "новый"))
            continue

        target_dir.mkdir(parents=True, exist_ok=True)
        if replaced:
            target_file.unlink()
        shutil.move(str(file), str(target_file))
        transferred.append((file.name, str(target_file.relative_to(DEST)), "ЗАМЕНА" if replaced else "новый"))

    if args.dry_run:
        print(json.dumps({
            "status": "dry-run",
            "would_transfer": len(transferred),
            "would_skip": len(issues),
            "transfers": transferred,
            "issues": issues,
        }, ensure_ascii=False, indent=2))
        return

    # --- Логи ---
    timestamp = datetime.now().isoformat(timespec="seconds")
    LOG_DIR.mkdir(parents=True, exist_ok=True)

    write_header = not TRANSFER_LOG.exists()
    with open(TRANSFER_LOG, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if write_header:
            writer.writerow(["Timestamp", "FileName", "TargetPath", "Action"])
        for name, path, action in transferred:
            writer.writerow([timestamp, name, path, action])

    with open(ISSUES_LOG, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["FileName", "Reason"])
        for name, reason in issues:
            writer.writerow([name, reason])

    print(json.dumps({
        "status": "ok",
        "transferred_count": len(transferred),
        "issues_count": len(issues),
        "issues": issues,
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
