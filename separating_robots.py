import pandas as pd
from pathlib import Path
from datetime import datetime
import logging

logger = logging.getLogger(__name__)

ROBOT_PATTERNS = [
    {"name": "smart_voice_bot", "pattern": "Smart Voice Bot Predictive"},
    {"name": "voice_bot", "pattern": "Voice Bot Consultation"},
    {"name": "reactivation_bot", "pattern": "Reactivation Bot"},
    {"name": "sales_bot", "pattern": "Sales Bot"},
]


def split_report_by_robot(input_path: str, output_dir: str = "split") -> list:
    """
    Разделяет общий Excel-отчёт на отдельные файлы по ботам
    на основе колонки 'Имя задания'.

    Args:
        input_path: путь к исходному Excel-файлу
        output_dir: директория для сохранения файлов (по умолчанию 'split')

    Returns:
        Список путей к созданным файлам.

    Raises:
        FileNotFoundError: если исходный файл не найден.
        ValueError: если в файле нет колонки 'Имя задания'.
    """
    input_path = Path(input_path)
    if not input_path.exists():
        raise FileNotFoundError(f"Файл не найден: {input_path}")

    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)

    df = pd.read_excel(input_path)

    if "Имя задания" not in df.columns:
        raise ValueError("В файле отсутствует колонка 'Имя задания'")

    robots_data = {robot["name"]: [] for robot in ROBOT_PATTERNS}

    for _, row in df.iterrows():
        task_name = str(row["Имя задания"]) if pd.notna(row["Имя задания"]) else ""
        for robot in ROBOT_PATTERNS:
            if robot["pattern"] in task_name:
                robots_data[robot["name"]].append(row)
                break  # предотвращает попадание в другие шаблоны

    created_files = []
    now = datetime.now()
    date_str = now.strftime("%d.%m.%Y")
    time_str = now.strftime("%H-%M")

    for robot_name, rows in robots_data.items():
        if not rows:
            logger.info(f"Для робота {robot_name} строк не найдено, файл не создаётся.")
            continue

        robot_df = pd.DataFrame(rows)
        filename = f"{robot_name}_{date_str}_{time_str}.xlsx"
        filepath = output_path / filename
        robot_df.to_excel(filepath, index=False)
        created_files.append(str(filepath))
        logger.info(f"Создан файл: {filepath} (строк: {len(robot_df)})")

    if not created_files:
        logger.warning("Не найдено строк ни для одного робота.")

    return created_files
