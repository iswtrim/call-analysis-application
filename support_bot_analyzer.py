import json
import logging
import re
import time
from datetime import datetime
from pathlib import Path

import openai
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from analyzer import APIKeyExpiredError, check_api_key_error, clean_json_response
from config import Config
from pivot_generator import filter_test_phones

SUPPORT_AGENT_NAME = "Inbound support"
CATEGORY_COLUMN = "Категория"
ERROR_TYPE_COLUMN = "Тип ошибки"
DIALOGUE_COLUMN = "Текст разговора с роботом"
RESULT_COLUMN = "results-confirmation"
BATCH_SIZE = 25


def build_dynamic_support_summary_dataframe(df):
    """
    Создаёт сводную таблицу с Excel-формулами (разделитель - запятая).
    Всего звонков = СЧЁТЗ(...) - 1 (убираем заголовок).
    Процент сбросов = сброс / всего (с защитой от деления на ноль).
    Убраны строки с уникальными номерами.
    """
    if df is None:
        return pd.DataFrame({"Показатель": [], "Значение": []})

    columns = list(df.columns)
    category_col_idx = columns.index(CATEGORY_COLUMN) + 1 if CATEGORY_COLUMN in columns else None

    if category_col_idx is None:
        return pd.DataFrame({"Показатель": ["Ошибка"], "Значение": ["Нет колонки Категория"]})

    category_col_letter = get_column_letter(category_col_idx)
    sheet_name = "Категоризация"
    data_range = f"'{sheet_name}'!{category_col_letter}:{category_col_letter}"

    # Формулы с запятой как разделителем
    total_formula = f"=СЧЁТЗ({data_range})-1"  # минус заголовок
    reset_formula = f"=СЧЁТЕСЛИ({data_range},\"сброс\")"
    operator_formula = f"=СЧЁТЕСЛИ({data_range},\"оператор\")"
    questions_formula = (
        f"=СЧЁТЕСЛИ({data_range},\"корректно\")"
        f"+СЧЁТЕСЛИ({data_range},\"не_корректно\")"
        f"+СЧЁТЕСЛИ({data_range},\"передали\")"
    )
    transferred_formula = f"=СЧЁТЕСЛИ({data_range},\"передали\")"
    correct_formula = (
        f"=СЧЁТЕСЛИ({data_range},\"корректно\")"
        f"+СЧЁТЕСЛИ({data_range},\"передали\")"
    )
    self_answer_formula = f"=СЧЁТЕСЛИ({data_range},\"корректно\")"
    called_back_formula = f"=СЧЁТЕСЛИ({data_range},\"звонили\")"
    negative_formula = f"=СЧЁТЕСЛИ({data_range},\"негатив\")"
    silence_formula = f"=СЧЁТЕСЛИ({data_range},\"алло\")"
    repeat_call_formula = f"=СЧЁТЕСЛИ({data_range},\"повторный_звонок\")"

    # Процент сбросов: если всего > 0, то сброс/всего, иначе пусто
    percent_formula = "=ЕСЛИ(B2>0,(B4/B2)*100,0)"  # B3 - всего, B4 - сброс (порядок строк фиксирован)
    percent_formula_unic = "=ЕСЛИ(B3>0,(B6/B3)*100,0)"  # B3 - всего, B4 - сброс (порядок строк фиксирован)

    summary_rows = [
        ("Всего звонков", total_formula),
        ("Количество уникальных номеров", 0),
        ("Сбросили на приветствии", reset_formula),
        ("Процент сбросов", percent_formula),
        ("Сбросили на приветствии (уникальные)", 0),
        ("Процент сбросов (уникальные)", 0),
        ("Попросили сотрудника", operator_formula),
        ("Задали вопросов", questions_formula),
        ("Робот передал информацию специалисту", transferred_formula),
        ("Корректная обработка вопроса", correct_formula),
        ("Робот самостоятельно ответил на вопрос", self_answer_formula),
        ("Вы мне звонили", called_back_formula),
        ("Негатив", negative_formula),
        ("Молчание", silence_formula),
        ("Повторный звонок", repeat_call_formula),
    ]

    return pd.DataFrame({"Показатель": [row[0] for row in summary_rows], "Значение": [row[1] for row in summary_rows]})


def _style_excel_sheet(ws, headers):
    """Применяет базовое форматирование к листу Excel."""
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment

    for col_idx in range(1, len(headers) + 1):
        max_len = 0
        for row in range(1, ws.max_row + 1):
            cell_value = ws.cell(row=row, column=col_idx).value
            if cell_value is not None:
                max_len = max(max_len, len(str(cell_value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 50)

    for row in range(1, ws.max_row + 1):
        for col in range(1, len(headers) + 1):
            ws.cell(row=row, column=col).border = thin_border


def _move_support_columns(df):
    """Переставляет колонки Категория и Тип ошибки после колонки Текст разговора с роботом."""
    if df is None:
        return df

    columns = list(df.columns)
    for column_name in [CATEGORY_COLUMN, ERROR_TYPE_COLUMN]:
        if column_name in columns:
            columns.remove(column_name)

    if DIALOGUE_COLUMN in columns:
        dialogue_index = columns.index(DIALOGUE_COLUMN)
        columns = columns[: dialogue_index + 1] + [CATEGORY_COLUMN, ERROR_TYPE_COLUMN] + columns[dialogue_index + 1:]
    else:
        columns = [CATEGORY_COLUMN, ERROR_TYPE_COLUMN] + columns

    return df.loc[:, columns]


def _build_unique_calls_sheet(df):
    """Создаёт лист с уникальными звонками по вызывающему номеру."""
    if df is None or df.empty:
        return pd.DataFrame()

    if "Вызывающий номер" not in df.columns:
        return pd.DataFrame()

    unique_df = df.dropna(subset=["Вызывающий номер"]).copy()
    if unique_df.empty:
        return pd.DataFrame()

    unique_df = unique_df.sort_values(by=["Вызывающий номер", "Дата начала звонка"], ascending=[True, False],
                                      na_position="last")
    unique_df = unique_df.drop_duplicates(subset=["Вызывающий номер"], keep="first")
    return unique_df


def save_support_report_with_summary(df, output_path, logger=None):
    """
    Сохраняет Excel-файл с листами 'Категоризация' и 'Сводка по категориям'.
    Используется openpyxl напрямую для корректной записи формул.
    Лист 'Уникальные звонки' НЕ создаётся.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    Path(output_path).parent.mkdir(exist_ok=True)

    # Переставляем колонки
    ordered_df = _move_support_columns(df)
    summary_df = build_dynamic_support_summary_dataframe(ordered_df)

    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])

    # ---- Лист 1: Категоризация ----
    ws1 = wb.create_sheet("Категоризация")
    for col_idx, col_name in enumerate(ordered_df.columns, 1):
        ws1.cell(row=1, column=col_idx, value=col_name)
    for row_idx, row_data in enumerate(ordered_df.itertuples(index=False), 2):
        for col_idx, value in enumerate(row_data, 1):
            ws1.cell(row=row_idx, column=col_idx, value=value)

    # ---- Лист 2: Сводка по категориям (с формулами) ----
    ws2 = wb.create_sheet("Сводка по категориям")
    for col_idx, col_name in enumerate(summary_df.columns, 1):
        ws2.cell(row=1, column=col_idx, value=col_name)
    for row_idx, row_data in enumerate(summary_df.itertuples(index=False), 2):
        ws2.cell(row=row_idx, column=1, value=row_data[0])  # Показатель
        val = row_data[1]
        if isinstance(val, str) and val.startswith("="):
            ws2.cell(row=row_idx, column=2, value=val)  # Формула
        else:
            ws2.cell(row=row_idx, column=2, value=val)  # Число/текст

    # ---- Стилизация всех листов ----
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        header_alignment = Alignment(horizontal="center", vertical="center")

        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment

        for col in range(1, ws.max_column + 1):
            max_len = 0
            for row in range(1, ws.max_row + 1):
                cell_value = ws.cell(row=row, column=col).value
                if cell_value is not None:
                    max_len = max(max_len, len(str(cell_value)))
            ws.column_dimensions[get_column_letter(col)].width = min(max_len + 2, 50)

        for row in range(1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).border = thin_border

    wb.save(output_path)

    if logger:
        logger.info(f"Сохранён Excel-файл: {output_path}")
        logger.info("Добавлен лист 'Сводка по категориям' с итоговой сводкой")


def extract_result_value(result, *field_names):
    """Извлекает значение из словаря ответа агента по нескольким возможным именам полей."""
    if not isinstance(result, dict):
        return ""

    for field_name in field_names:
        if field_name in result and result[field_name] not in (None, ""):
            value = str(result[field_name]).strip()
            if value:
                return value

    return ""


def create_batch_input(batch_data):
    """Создаёт входные данные для пакета диалогов."""
    input_text = ""
    for i, data in enumerate(batch_data, 1):
        input_text += f"ДИАЛОГ {i}\n"
        input_text += f"Текст разговора: {data['dialogue']}\n"
        input_text += f"Результат диалога (Twin): {data['result']}\n\n"
    return input_text


def parse_agent_response(response_text, expected_count, logger):
    """Парсит JSON-ответ агента с категориями."""
    cleaned_text = clean_json_response(response_text)

    if check_api_key_error(cleaned_text):
        raise APIKeyExpiredError(f"API-ключ истек (в очищенном ответе): {cleaned_text[:200]}...")

    logger.info(f"Длина ответа от нейросети: {len(cleaned_text)} символов")
    logger.info(f"Первые 300 символов ответа: {cleaned_text[:300]}...")

    try:
        parsed = json.loads(cleaned_text)
    except json.JSONDecodeError as exc:
        json_match = re.search(r"\[.*\]", cleaned_text, re.DOTALL)
        if not json_match:
            raise exc
        parsed = json.loads(json_match.group(0))

    if isinstance(parsed, list):
        results = parsed
    elif isinstance(parsed, dict):
        if isinstance(parsed.get("results"), list):
            results = parsed["results"]
        elif CATEGORY_COLUMN in parsed or ERROR_TYPE_COLUMN in parsed:
            results = [parsed]
        else:
            list_value = next((value for value in parsed.values() if isinstance(value, list)), None)
            if list_value is None:
                raise ValueError(
                    f"Нейросеть вернула непонятный формат. Ключи: {list(parsed.keys())}"
                )
            results = list_value
    else:
        raise ValueError(f"Ответ нейросети имеет неподдерживаемый тип: {type(parsed)}")

    if len(results) != expected_count:
        logger.warning(
            f"Агент вернул {len(results)} результатов, ожидалось {expected_count}"
        )
        if len(results) < expected_count:
            for _ in range(expected_count - len(results)):
                results.append({
                    CATEGORY_COLUMN: "Ошибка",
                    ERROR_TYPE_COLUMN: "Неполный ответ от нейросети",
                })

    return results[:expected_count]


def analyze_support_batch(batch_data, logger, client, prompt_id):
    """Анализирует пакет диалогов через агента Inbound Support."""
    response_text = ""

    try:
        logger.info(f"Отправка пакета из {len(batch_data)} диалогов...")
        start_batch_time = time.time()

        response = client.responses.create(
            prompt={"id": prompt_id},
            input=create_batch_input(batch_data),
        )

        batch_time = time.time() - start_batch_time
        logger.info(f"Время обработки пачки: {batch_time:.2f} сек")

        if not hasattr(response, "output_text") or not response.output_text:
            raise ValueError("Пустой ответ от агента")

        response_text = response.output_text

        if check_api_key_error(response_text):
            raise APIKeyExpiredError(
                f"API-ключ истек (в ответе агента): {response_text[:200]}..."
            )

        results = parse_agent_response(response_text, len(batch_data), logger)

        batch_results = []
        for i, (result, call_data) in enumerate(zip(results, batch_data)):
            if isinstance(result, dict):
                category = str(extract_result_value(result, CATEGORY_COLUMN, "category", "Категория")).strip()
                error_type = extract_result_value(result, ERROR_TYPE_COLUMN, "Тип_ошибки", "error_type", "type_error")
            else:
                category = "Ошибка"
                error_type = f"Неверный формат элемента: {type(result)}"

            if not category:
                category = "Ошибка"
                error_type = error_type or "Пустая категория в ответе агента"

            normalized_category = category.lower()
            if normalized_category in {"корректно", "сброс"}:
                error_type = ""
            elif normalized_category in {"не_корректно", "некорректно", "не корректно"}:
                error_type = error_type or "Не указана причина ошибки"
            elif not error_type:
                error_type = ""

            logger.info(
                f"  Диалог {i + 1} (ID: {call_data['call_id']}): "
                f"{CATEGORY_COLUMN}='{category}', {ERROR_TYPE_COLUMN}='{error_type}'"
            )

            batch_results.append({
                "df_index": call_data["df_index"],
                "category": category,
                "error_type": error_type,
            })

        return batch_results

    except openai.AuthenticationError as exc:
        raise APIKeyExpiredError(
            f"API-ключ недействителен (AuthenticationError): {str(exc)[:200]}..."
        ) from exc

    except openai.APIError as exc:
        error_msg = str(exc)
        logger.error(f"Ошибка API OpenAI: {error_msg}")

        if check_api_key_error(error_msg):
            raise APIKeyExpiredError(f"API-ключ истек (APIError): {error_msg[:200]}...") from exc

        return [
            {
                "df_index": data["df_index"],
                "category": "Ошибка API",
                "error_type": f"Ошибка API: {error_msg[:100]}",
            }
            for data in batch_data
        ]

    except APIKeyExpiredError:
        raise

    except Exception as exc:
        error_msg = str(exc)
        logger.error(f"Общая ошибка при анализе пачки: {error_msg}")

        if check_api_key_error(error_msg):
            raise APIKeyExpiredError(f"API-ключ истек: {error_msg[:200]}...") from exc

        if response_text and check_api_key_error(response_text):
            raise APIKeyExpiredError(
                f"API-ключ истек (в ответе): {response_text[:200]}..."
            ) from exc

        return [
            {
                "df_index": data["df_index"],
                "category": "Ошибка API",
                "error_type": f"Ошибка при обработке: {error_msg[:100]}",
            }
            for data in batch_data
        ]


def analyze_support_report(input_excel_path):
    """
    Анализирует отчёт и добавляет колонки 'Категория' и 'Тип_ошибки' в конец файла.

    Возвращает: (DataFrame_с_результатами, путь_к_логу)
    """
    prompt_id = Config.AGENTS.get(SUPPORT_AGENT_NAME)
    if not prompt_id:
        raise ValueError(
            f"Агент '{SUPPORT_AGENT_NAME}' не найден. Проверьте YANDEX_AGENT_SUPPORT в .env"
        )

    timestamp = datetime.now().strftime("%d.%m.%Y_%H-%M")
    original_name = Path(input_excel_path).stem
    file_prefix = f"support_{timestamp}_{original_name}"
    log_filename = f"logs/{file_prefix}.log"
    error_filename = f"logs/ERROR_{file_prefix}.txt"

    Path("logs").mkdir(exist_ok=True)

    logger = logging.getLogger(f"support_analyzer_{timestamp}")
    logger.setLevel(logging.INFO)
    logger.propagate = False
    logger.handlers = []

    file_handler = logging.FileHandler(log_filename, encoding="utf-8")
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    )
    logger.addHandler(file_handler)

    logger.info("=" * 60)
    logger.info(f"НАЧАЛО АНАЛИЗА: {timestamp}")
    logger.info(f"Агент: {SUPPORT_AGENT_NAME}")
    logger.info(f"Prompt ID: {prompt_id[:10]}...")
    logger.info(f"Исходный файл: {input_excel_path}")
    logger.info(f"Лог-файл: {log_filename}")
    logger.info("=" * 60)

    try:
        logger.info(f"Загрузка файла: {input_excel_path}")
        df = pd.read_excel(input_excel_path)
        logger.info(f"Загружено {len(df)} строк, {len(df.columns)} колонок")

        df = filter_test_phones(df, logger)

        columns_to_remove = ["Идентификатор задания", "Имя задания"]
        existing_columns = [col for col in columns_to_remove if col in df.columns]
        if existing_columns:
            df = df.drop(columns=existing_columns)
            logger.info(f"Удалены колонки: {', '.join(existing_columns)}")

        required_cols = [DIALOGUE_COLUMN, RESULT_COLUMN]
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            error_msg = f"Отсутствуют колонки: {missing_cols}"
            logger.error(error_msg)
            raise ValueError(error_msg)

        logger.info("Все необходимые колонки присутствуют")

        if CATEGORY_COLUMN in df.columns:
            df = df.drop(columns=[CATEGORY_COLUMN])
        if ERROR_TYPE_COLUMN in df.columns:
            df = df.drop(columns=[ERROR_TYPE_COLUMN])

        df[CATEGORY_COLUMN] = ""
        df[ERROR_TYPE_COLUMN] = ""

        logger.info("Создаю клиент для Yandex Cloud...")
        client = openai.OpenAI(
            api_key=Config.YANDEX_API_KEY,
            base_url=Config.YANDEX_BASE_URL,
            project=Config.YANDEX_PROJECT_ID,
        )

        calls_to_process = []
        for idx, row in df.iterrows():
            dialogue = (
                str(row[DIALOGUE_COLUMN]) if pd.notna(row[DIALOGUE_COLUMN]) else ""
            )
            call_id = row.get("Идентификатор звонка", f"Строка_{idx + 2}")
            calls_to_process.append({
                "df_index": idx,
                "call_id": call_id,
                "dialogue": dialogue,
                "result": (
                    str(row[RESULT_COLUMN]) if pd.notna(row[RESULT_COLUMN]) else ""
                ),
            })

        logger.info(f"Подготовлено {len(calls_to_process)} диалогов для анализа")

        batches = [
            calls_to_process[i:i + BATCH_SIZE]
            for i in range(0, len(calls_to_process), BATCH_SIZE)
            if calls_to_process[i:i + BATCH_SIZE]
        ]
        logger.info(f"Создано {len(batches)} пачек по {BATCH_SIZE} диалогов")

        start_time = time.time()
        total_processed = 0

        for batch_num, batch in enumerate(batches, 1):
            logger.info(
                f"Обработка пачки {batch_num}/{len(batches)} ({len(batch)} диалогов)..."
            )

            batch_results = analyze_support_batch(batch, logger, client, prompt_id)

            for result in batch_results:
                df_index = result["df_index"]
                if 0 <= df_index < len(df):
                    df.at[df_index, CATEGORY_COLUMN] = result["category"]
                    df.at[df_index, ERROR_TYPE_COLUMN] = result["error_type"]
                else:
                    logger.error(f"Неверный индекс {df_index}")

            total_processed += len(batch)

            if batch_num < len(batches):
                time.sleep(1)

        category_counts = df[CATEGORY_COLUMN].value_counts()
        logger.info("=" * 60)
        logger.info("ИТОГИ АНАЛИЗА:")
        logger.info(f"   Всего диалогов: {len(df)}")
        logger.info(f"   Проанализировано: {total_processed}")
        logger.info(f"   Время обработки: {time.time() - start_time:.1f} секунд")
        logger.info("   Распределение категорий:")
        for category, count in category_counts.items():
            logger.info(f"     {category}: {count}")
        logger.info(f"   Лог-файл: {log_filename}")
        logger.info("=" * 60)

        return df, log_filename

    except APIKeyExpiredError:
        logger.critical("=" * 60)
        logger.critical("АНАЛИЗ ПРЕРВАН ИЗ-ЗА ИСТЕЧЕНИЯ API-КЛЮЧА!")
        logger.critical("=" * 60)
        raise

    except Exception as exc:
        logger.error("КРИТИЧЕСКАЯ ОШИБКА:")
        logger.error(f"   Тип ошибки: {type(exc).__name__}")
        logger.error(f"   Сообщение: {str(exc)}")

        with open(error_filename, "w", encoding="utf-8") as error_file:
            error_file.write(f"Ошибка при анализе {input_excel_path}:\n")
            error_file.write(f"Время: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
            error_file.write(f"Агент: {SUPPORT_AGENT_NAME}\n")
            error_file.write(f"Тип ошибки: {type(exc).__name__}\n")
            error_file.write(f"Сообщение: {str(exc)}\n\n")
            import traceback
            error_file.write("Трассировка:\n")
            error_file.write(traceback.format_exc())

        return None, error_filename
