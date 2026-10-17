import pandas as pd
from datetime import datetime
from pathlib import Path
import logging
from typing import Dict, Tuple

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

from config import Config

# Игнорируем предупреждения pandas
import warnings

warnings.filterwarnings('ignore')

logger = logging.getLogger(__name__)

# Успешные результаты (для подсветки зеленым), регистр важен!
SUCCESSFUL_RESULTS = [
    'Лид',
    'На оператора',
    'Согласие',
    'Получен контакт ЛПР',
    'Прислать информацию',
    'Успех',
    'Актуально'
]

# Неконтактные результаты (для расчёта контактности), регистр важен!
NON_CONTACT_RESULTS = {
    'Автоответчик',
    'Тишина',
    'Сброс',
    'Спам',
    'Умный автоответчик',
    'Попали в компанию',
    'Лид в браке'
}


def filter_test_phones(df, logger=None):
    """
    Удаляет строки, где номер телефона совпадает с тестовыми номерами
    из Config.TEST_PHONES. Для входящих отчётов используется колонка
    'Вызывающий номер', для исходящих — 'Вызываемый номер'.

    Args:
        df: DataFrame с данными отчёта
        logger: логгер для записи количества удалённых строк

    Returns:
        Отфильтрованный DataFrame без тестовых номеров.
    """
    if not Config.TEST_PHONES:
        return df

    # Определяем тип отчёта по наличию колонки 'Результат AMD'
    is_incoming = 'Результат AMD' not in df.columns

    # Выбираем правильную колонку
    if is_incoming:
        phone_col = 'Вызывающий номер'
    else:
        phone_col = 'Вызываемый номер'

    if phone_col not in df.columns:
        if logger:
            logger.warning(f"Колонка '{phone_col}' не найдена для фильтрации тестовых номеров.")
        return df

    initial_len = len(df)
    df_filtered = df[~df[phone_col].astype(str).isin(Config.TEST_PHONES)].copy()
    removed = initial_len - len(df_filtered)

    if removed > 0 and logger:
        logger.info(f"Удалено {removed} строк с тестовыми номерами из колонки '{phone_col}'.")

    return df_filtered


def is_incoming_report(df: pd.DataFrame) -> bool:
    """Возвращает True, если это входящий отчёт (отсутствует колонка 'Результат AMD')."""
    return 'Результат AMD' not in df.columns


def add_comment_to_cell(cell, formula_text, logger=None):
    """
    Добавляет комментарий с описанием формулы к ячейке Excel.
    Изменяет переданную ячейку на месте, ничего не возвращает.
    Длинные комментарии (>500 символов) обрезаются.

    Args:
        cell: ячейка Excel (openpyxl)
        formula_text: текст формулы/описания
        logger: логгер для записи предупреждений

    Returns:
        None
    """
    try:
        if not formula_text or pd.isna(formula_text):
            return

        # Ограничиваем длину комментария
        if len(str(formula_text)) > 500:
            formula_text = str(formula_text)[:497] + "..."

        comment = Comment(str(formula_text), 'Система')
        comment.width = 300
        comment.height = 100
        cell.comment = comment

    except Exception as e:
        if logger:
            logger.warning(f"Не удалось добавить комментарий к ячейке: {e}")
        else:
            # fallback на глобальный логгер
            logging.getLogger(__name__).warning(f"Не удалось добавить комментарий к ячейке: {e}")


def create_general_table(original_df, last_calls_df, logger, include_correct_calls=True, is_incoming=False):
    """
    Создает таблицу 'Общая' с метриками.

    Args:
        original_df: DataFrame исходного отчёта
        last_calls_df: DataFrame отчёта по последнему звонку
        logger: логгер
        include_correct_calls: True для полного отчёта, False для режима только сводные
        is_incoming: True для входящих звонков, False для исходящих

    Returns:
        Кортеж (DataFrame с метриками, список формул для комментариев)
    """
    try:
        logger.info("Создание таблицы 'Общая'...")

        table_data = []
        formulas = []  # Список формул для комментариев

        # 1. Количество кандидатов
        candidates_count = last_calls_df[
            'Идентификатор звонка'].nunique() if 'Идентификатор звонка' in last_calls_df.columns else len(last_calls_df)
        formula_candidates = "Количество 'Идентификатор звонка' в листе 'Отчёт по последнему звонку'"
        table_data.append({'Показатель': 'Количество кандидатов', 'Значение': candidates_count})
        formulas.append(formula_candidates)

        # 2. Количество звонков
        calls_count = original_df[
            'Идентификатор звонка'].nunique() if 'Идентификатор звонка' in original_df.columns else len(original_df)
        formula_calls = "Количество 'Идентификатор звонка' в листе 'Исходный отчёт'"
        table_data.append({'Показатель': 'Количество звонков', 'Значение': calls_count})
        formulas.append(formula_calls)

        # Пустая строка
        table_data.append({'Показатель': '', 'Значение': ''})
        formulas.append('')

        # 3. Контактность относительно всех звонков
        contact_calls = 0  # Будет рассчитано позже
        formula_contact_calls = "Итого из таблицы 'Результаты в исходном отчёте' - сумма автоответчиков, тишины и сбросов"
        table_data.append({'Показатель': 'Контактность относительно всех звонков', 'Значение': contact_calls})
        formulas.append(formula_contact_calls)

        # 4. Контактность относительно всех номеров
        contact_numbers = 0  # Будет рассчитано позже
        formula_contact_numbers = "Итого из таблицы 'Результаты в отчёте по последнему звонку' - сумма автоответчиков, тишины и сбросов"
        table_data.append({'Показатель': 'Контактность относительно всех номеров', 'Значение': contact_numbers})
        formulas.append(formula_contact_numbers)

        # Пустая строка
        table_data.append({'Показатель': '', 'Значение': ''})
        formulas.append('')

        # 5. Стоимость прозвона
        col_cost = 'O' if is_incoming else 'Q'
        call_cost_formula = f"=СУММ('Исходный отчёт'!{col_cost}:{col_cost})"
        formula_call_cost = "Сумма столбца 'Стоимость звонка' в листе 'Исходный отчёт'"
        table_data.append({'Показатель': 'Стоимость прозвона', 'Значение': call_cost_formula})
        formulas.append(formula_call_cost)

        # 6. CPL (стоимость лида)
        # Для формул, завязанных на исходный отчёт, берём успешные результаты из original_df
        if 'results-confirmation' in original_df.columns:
            successful_results = original_df['results-confirmation'].isin(SUCCESSFUL_RESULTS)
            successful_count = original_df[successful_results].shape[0]
        else:
            successful_count = 0

        # Формула CPL в зависимости от режима
        if include_correct_calls:
            # Полный отчёт - формула с учётом лидов и фродов
            cpl_formula = '=J10/(J14-J15)'
            formula_desc = "Стоимость прозвона / (Всего лидов - Количество ошибочных лидов)"
        else:
            # Только сводные - формула ссылается на ячейку со стоимостью
            if successful_count > 0:
                cpl_formula = f'=J10/{successful_count}'
            else:
                cpl_formula = '=0'
            formula_desc = "Стоимость прозвона / Количество успешных результатов"

        table_data.append({'Показатель': 'CPL (стоимость лида)', 'Значение': cpl_formula})
        formulas.append(formula_desc)

        # Пустая строка
        table_data.append({'Показатель': '', 'Значение': ''})
        formulas.append('')

        # Показатели ТОЛЬКО ДЛЯ ПОЛНОГО ОТЧЁТА:
        # 7. Количество корректных звонков
        if include_correct_calls:
            excel_formula_correct_calls = "=B14-N5"
            table_data.append({'Показатель': 'Количество корректных звонков', 'Значение': excel_formula_correct_calls})
            formulas.append("Количество корректных звонков")

            # 8. Всего лидов
            table_data.append({'Показатель': 'Всего лидов', 'Значение': successful_count})
            formulas.append("Количество успешных результатов в исходном отчёте")

            # 9. Количество ошибочных лидов
            col_err = 'T' if is_incoming else 'V'  # колонка 'Корректность' в листе 'Отчёт по ошибкам'
            col_res = 'V' if is_incoming else 'X'  # колонка 'results-confirmation'
            terms = []
            for res in SUCCESSFUL_RESULTS:
                terms.append(
                    f'COUNTIFS(\'Отчёт по ошибкам\'!{col_err}:{col_err}, "Ошибка", \'Отчёт по ошибкам\'!{col_res}:{col_res}, "{res}") + '
                    f'COUNTIFS(\'Отчёт по ошибкам\'!{col_err}:{col_err}, "Ошибка распознавания", \'Отчёт по ошибкам\'!{col_res}:{col_res}, "{res}")'
                )
            fraud_formula = '=' + '+'.join(terms)
            table_data.append({'Показатель': 'Количество ошибочных лидов', 'Значение': fraud_formula})
            formulas.append("Количество лидов, которые оказались некорректными (фрод)")

        # Создаем DataFrame
        general_df = pd.DataFrame(table_data)

        # Добавляем колонку с процентами (значения будут заполнены позже)
        percentages = [''] * len(general_df)
        general_df['Процент'] = percentages

        logger.info(f"✅ Таблица 'Общая' создана: {len(general_df)} строк")
        logger.info(f"   Включая корректные звонки: {include_correct_calls}")

        return general_df, formulas

    except Exception as e:
        logger.error(f"Ошибка при создании таблицы 'Общая': {e}")

        import traceback
        logger.error(traceback.format_exc())

        return pd.DataFrame(), []


def create_full_report(original_path, errors_df, agent_name, logger):
    """
    Создаёт полный отчёт с 5 листами: Сводная, Сводная телефонии,
    Отчёт по ошибкам, Отчёт по последнему звонку, Исходный отчёт.

    Args:
        original_path: путь к исходному Excel-файлу с отчётом
        errors_df: DataFrame с найденными ошибками (из AI-анализа)
        agent_name: имя агента для AI-анализа
        logger: логгер для записи хода выполнения

    Returns:
        Путь к сохранённому файлу отчёта или None при ошибке.
    """
    import pivot_generator
    pivot_generator.logger = logger

    try:
        logger.info(f"Создание полного отчёта для агента {agent_name}...")

        # 1. Загружаем исходный файл
        original_df = pd.read_excel(original_path)
        logger.info(f"✅ Загружен исходный файл: {len(original_df)} строк")

        original_df = filter_test_phones(original_df, logger)  # УДАЛЯЕМ ТЕСТОВЫЕ НОМЕРА

        # Проверяем тип отчёта (входящие или исходящие звонки)
        is_incoming = is_incoming_report(original_df)

        # Удаляем ненужные колонки
        columns_to_remove = ['Идентификатор задания', 'Имя задания']
        existing_columns = [col for col in columns_to_remove if col in original_df.columns]

        if existing_columns:
            original_df = original_df.drop(columns=existing_columns)
            logger.info(f"✅ Удалены колонки: {', '.join(existing_columns)}")
            logger.info(f"   Осталось колонок: {len(original_df.columns)}")

        # 2. Создаем все необходимые таблицы
        logger.info("Создание сводных таблиц...")

        # Существующие таблицы
        status_pivot = create_telephony_status_pivot(original_df, logger)

        if not is_incoming:
            auto_pivot = create_auto_answer_pivot(original_df, logger)
        else:
            auto_pivot = pd.DataFrame()  # пустой DataFrame, чтобы не было ошибок

        last_calls_df = create_last_calls_report(original_df, logger, is_incoming=is_incoming)
        original_results_summary = create_results_summary(original_df, "Исходный отчёт", logger)
        last_calls_results_summary = create_results_summary(last_calls_df, "Отчёт по последнему звонку", logger)

        # Новые таблицы
        general_df, general_formulas = create_general_table(
            original_df,
            last_calls_df,
            logger,
            include_correct_calls=True,
            is_incoming=is_incoming
        )

        correctness_df, _, correctness_comments = create_correctness_summary(logger, is_incoming=is_incoming)
        # Обновляем значения контактности в general_df
        if not general_df.empty:
            contact_calls_idx = general_df[general_df['Показатель'] == 'Контактность относительно всех звонков'].index
            contact_numbers_idx = general_df[general_df['Показатель'] == 'Контактность относительно всех номеров'].index

            if len(contact_calls_idx) > 0:
                contact_calls = calculate_contactness(original_results_summary)
                general_df.at[contact_calls_idx[0], 'Значение'] = contact_calls
                calls_count = general_df.at[1, 'Значение']
                if calls_count > 0:
                    general_df.at[contact_calls_idx[0], 'Процент'] = contact_calls / calls_count

            if len(contact_numbers_idx) > 0:
                contact_numbers = calculate_contactness(last_calls_results_summary)
                general_df.at[contact_numbers_idx[0], 'Значение'] = contact_numbers
                candidates_count = general_df.at[0, 'Значение']
                if candidates_count > 0:
                    general_df.at[contact_numbers_idx[0], 'Процент'] = contact_numbers / candidates_count

        # 4. Подготовка к сохранению
        timestamp = datetime.now().strftime('%d.%m.%Y_%H-%M')
        original_name = Path(original_path).stem
        agent_filename = agent_name.replace(" ", "_").replace("-", "_").lower()

        processed_dir = "processed"
        Path(processed_dir).mkdir(exist_ok=True)

        output_filename = f"{processed_dir}/FULL_{agent_filename}_{timestamp}_{original_name}.xlsx"

        # 5. Создаем книгу Excel
        wb = Workbook()
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])

        # ===== 1. ЛИСТ "Сводная" =====
        ws_summary = wb.create_sheet(title='Сводная')

        # РАЗМЕЩЕНИЕ 4 ТАБЛИЦ В СТРОКУ:
        # A-C: Результаты в исходном отчёте
        # D: ПУСТОЙ
        # E-G: Результаты в отчёте по последнему звонку
        # H: ПУСТОЙ
        # I-K: Общая
        # L: ПУСТОЙ
        # M-O: Сводная корректности

        # ТАБЛИЦА 1: Результаты в исходном отчёте
        row = 1
        start_col = 1
        ws_summary.cell(row=row, column=start_col, value='Результаты в исходном отчёте').font = Font(bold=True, size=12)
        row += 2

        if not original_results_summary.empty and 'Ошибка' not in original_results_summary.columns:
            headers = list(original_results_summary.columns)
            for col_idx, header in enumerate(headers, 0):
                ws_summary.cell(row=row, column=start_col + col_idx, value=header)

            table1_start_row = row
            row += 1

            for _, data_row in original_results_summary.iterrows():
                for col_idx, value in enumerate(data_row.tolist(), 0):
                    ws_summary.cell(row=row, column=start_col + col_idx, value=value)
                row += 1

            format_excel_table(ws_summary, table1_start_row, start_col=start_col, is_results_table=True, logger=logger)

        # ПУСТОЙ СТОЛБЕЦ D

        # ТАБЛИЦА 2: Результаты в отчёте по последнему звонку
        row = 1
        start_col = 5
        ws_summary.cell(row=row, column=start_col, value='Результаты в отчёте по последнему звонку').font = Font(
            bold=True, size=12)
        row += 2

        if not last_calls_results_summary.empty and 'Ошибка' not in last_calls_results_summary.columns:
            headers = list(last_calls_results_summary.columns)
            for col_idx, header in enumerate(headers, 0):
                ws_summary.cell(row=row, column=start_col + col_idx, value=header)

            table2_start_row = row
            row += 1

            for _, data_row in last_calls_results_summary.iterrows():
                for col_idx, value in enumerate(data_row.tolist(), 0):
                    ws_summary.cell(row=row, column=start_col + col_idx, value=value)
                row += 1

            format_excel_table(ws_summary, table2_start_row, start_col=start_col, is_results_table=True, logger=logger)

        # ПУСТОЙ СТОЛБЕЦ H

        # ТАБЛИЦА 3: Общая
        row = 1
        start_col = 9
        ws_summary.cell(row=row, column=start_col, value='Общая').font = Font(bold=True, size=12)
        row += 2

        if not general_df.empty:
            headers = ['Показатель', 'Значение', 'Процент']
            for col_idx, header in enumerate(headers, 0):
                ws_summary.cell(row=row, column=start_col + col_idx, value=header)

            table3_start_row = row
            row += 1

            for idx, data_row in general_df.iterrows():
                # Показатель
                ws_summary.cell(row=row, column=start_col, value=data_row['Показатель'])

                # ЗНАЧЕНИЕ - сюда пишем либо формулу, либо число
                value = data_row['Значение']
                value_cell = ws_summary.cell(row=row, column=start_col + 1)

                if isinstance(value, str) and value.startswith('='):
                    value_cell.value = value  # формула Excel
                else:
                    value_cell.value = value  # обычное число

                if idx < len(general_formulas) and general_formulas[idx]:
                    add_comment_to_cell(value_cell, general_formulas[idx], logger=logger)

                # ПРОЦЕНТ - только для показателей, где он нужен
                percent_cell = ws_summary.cell(row=row, column=start_col + 2)
                indicator = data_row['Показатель']

                if indicator == 'Контактность относительно всех звонков':
                    percent_cell.value = '=J7/J5'
                    percent_cell.number_format = '0.00%'
                    add_comment_to_cell(percent_cell, "Контактность относительно всех звонков / Количество звонков",
                                        logger=logger)

                elif indicator == 'Контактность относительно всех номеров':
                    percent_cell.value = '=J8/J4'
                    percent_cell.number_format = '0.00%'
                    add_comment_to_cell(percent_cell, "Контактность относительно всех номеров / Количество кандидатов",
                                        logger=logger)

                elif indicator == 'Количество корректных звонков':
                    percent_cell.value = '=1 - N5 / B14'
                    percent_cell.number_format = '0.00%'
                    add_comment_to_cell(
                        percent_cell,
                        "1 - Количество ошибок / Количество звонков с каким-либо результатом",
                        logger=logger,
                    )

                else:
                    # Для всех остальных строк (CPL, Всего лидов, Количество ошибочных лидов и т.д.)
                    # процент не нужен - оставляем пустым
                    if pd.notna(data_row['Процент']) and data_row['Процент'] != '':
                        percent_cell.value = data_row['Процент']
                        if isinstance(data_row['Процент'], (int, float)):
                            percent_cell.number_format = '0.00%'

                row += 1

            format_excel_table(ws_summary, table3_start_row, start_col=start_col, is_results_table=False, logger=logger)

        # ПУСТОЙ СТОЛБЕЦ L

        # ТАБЛИЦА 4: Сводная корректности
        row = 1
        start_col = 13
        ws_summary.cell(row=row, column=start_col, value='Сводная корректности').font = Font(bold=True, size=12)
        row += 2

        if not correctness_df.empty:
            headers = ['Корректность', 'Количество']
            for col_idx, header in enumerate(headers, 0):
                ws_summary.cell(row=row, column=start_col + col_idx, value=header)

            table4_start_row = row
            row += 1

            for idx, data_row in correctness_df.iterrows():
                ws_summary.cell(row=row, column=start_col, value=data_row['Корректность'])

                formula_excel = data_row['Количество']
                cell = ws_summary.cell(row=row, column=start_col + 1)

                if isinstance(formula_excel, str) and formula_excel.startswith('='):
                    cell.value = formula_excel
                    if cell.value and str(cell.value).startswith('=@'):
                        cell.value = str(cell.value).replace('=@', '=')
                else:
                    cell.value = formula_excel

                if idx < len(correctness_comments) and correctness_comments[idx]:
                    add_comment_to_cell(cell, correctness_comments[idx], logger=logger)

                row += 1

            format_excel_table(ws_summary, table4_start_row, start_col=start_col, is_results_table=False, logger=logger)
            ws_summary.column_dimensions[
                'N'].width = 15  # Устанавливаем ширину колонки Количество в Сводной корректности

        logger.info(f"✅ Создан лист 'Сводная' с 4 таблицами")

        # ===== 2. ЛИСТ "Сводная телефонии" =====
        ws_telephony = wb.create_sheet(title='Сводная телефонии')

        # Таблица 1: Сводная по статусу телефонии
        row = 1
        ws_telephony.cell(row=row, column=1, value='Сводная по статусу телефонии').font = Font(bold=True, size=12)
        row += 2
        if not status_pivot.empty and 'Ошибка' not in status_pivot.columns:
            headers = list(status_pivot.columns)
            for col_idx, header in enumerate(headers, 1):
                ws_telephony.cell(row=row, column=col_idx, value=header)
            start_row = row
            row += 1
            for _, data_row in status_pivot.iterrows():
                for col_idx, value in enumerate(data_row.tolist(), 1):
                    ws_telephony.cell(row=row, column=col_idx, value=value)
                row += 1
            format_excel_table(ws_telephony, start_row, is_results_table=False, logger=logger)

        # Таблица 2: Сводная по детектору автоответчиков – только для исходящих
        if not is_incoming:
            row += 3
            ws_telephony.cell(row=row, column=1, value='Сводная по детектору автоответчиков').font = Font(bold=True,
                                                                                                          size=12)
            row += 2
            if not auto_pivot.empty and 'Ошибка' not in auto_pivot.columns:
                headers = list(auto_pivot.columns)
                for col_idx, header in enumerate(headers, 1):
                    ws_telephony.cell(row=row, column=col_idx, value=header)
                start_row = row
                row += 1
                for _, data_row in auto_pivot.iterrows():
                    for col_idx, value in enumerate(data_row.tolist(), 1):
                        ws_telephony.cell(row=row, column=col_idx, value=value)
                    row += 1
                format_excel_table(ws_telephony, start_row, is_results_table=False, logger=logger)

        logger.info(f"✅ Создан лист 'Сводная телефонии'")

        # ===== 3. ЛИСТ "Отчёт по ошибкам" =====
        ws_errors = wb.create_sheet(title='Отчёт по ошибкам')

        if not errors_df.empty:
            headers = list(errors_df.columns)
            for col_idx, header in enumerate(headers, 1):
                ws_errors.cell(row=1, column=col_idx, value=header)

            for row_idx, (_, data_row) in enumerate(errors_df.iterrows(), 2):
                for col_idx, value in enumerate(data_row.tolist(), 1):
                    ws_errors.cell(row=row_idx, column=col_idx, value=value)

            format_excel_table(ws_errors, 1, logger=logger)
            logger.info(f"✅ Создан лист 'Отчёт по ошибкам' ({len(errors_df)} строк)")
        else:
            ws_errors.cell(row=1, column=1, value='Нет данных по ошибкам')
            logger.info("⚠️ Лист 'Отчёт по ошибкам' пуст")

        # ===== 4. ЛИСТ "Отчёт по последнему звонку" =====
        if not last_calls_df.empty:
            ws_last_calls = wb.create_sheet(title='Отчёт по последнему звонку')

            ws_last_calls.cell(row=1, column=1, value='Отчёт по последнему звонку').font = Font(bold=True, size=14)
            ws_last_calls.cell(row=2, column=1, value=f'Всего уникальных номеров: {len(last_calls_df)}')

            row = 4
            headers = list(last_calls_df.columns)
            for col_idx, header in enumerate(headers, 1):
                ws_last_calls.cell(row=row, column=col_idx, value=header)

            start_row = row
            row += 1

            for _, data_row in last_calls_df.iterrows():
                for col_idx, value in enumerate(data_row.tolist(), 1):
                    ws_last_calls.cell(row=row, column=col_idx, value=value)
                row += 1

            add_repeat_count_column(ws_last_calls, is_incoming, logger)

            format_excel_table(ws_last_calls, start_row, logger=logger)

            logger.info(f"✅ Создан лист 'Отчёт по последнему звонку' ({len(last_calls_df)} строк)")
        else:
            logger.info("⚠️ Лист 'Отчёт по последнему звонку' не создан (нет данных)")

        # ===== 5. ЛИСТ "Исходный отчёт" =====
        ws_original = wb.create_sheet(title='Исходный отчёт')

        headers = list(original_df.columns)
        for col_idx, header in enumerate(headers, 1):
            ws_original.cell(row=1, column=col_idx, value=header)

        for row_idx, (_, data_row) in enumerate(original_df.iterrows(), 2):
            for col_idx, value in enumerate(data_row.tolist(), 1):
                ws_original.cell(row=row_idx, column=col_idx, value=value)

        format_excel_table(ws_original, 1, logger=logger)
        logger.info(f"✅ Создан лист 'Исходный отчёт' ({len(original_df)} строк)")

        # 6. Сохраняем файл
        wb.save(output_filename)
        logger.info(f"✅ Файл сохранен: {output_filename}")
        logger.info(f"   Листы в файле: {wb.sheetnames}")

        return output_filename

    except Exception as e:
        logger.error(f"❌ Ошибка при создании полного отчёта: {e}")
        import traceback
        logger.error(traceback.format_exc())

        # Создаем файл с ошибкой
        error_filename = f"logs/ERROR_FULL_REPORT_{agent_filename}_{timestamp}.txt"
        with open(error_filename, 'w', encoding='utf-8') as f:
            f.write(f"Ошибка при создании полного отчёта:\n")
            f.write(f"Время: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
            f.write(f"Агент: {agent_name}\n")
            f.write(f"Файл: {original_path}\n")
            f.write(f"Ошибка: {str(e)}\n\n")
            f.write("Трассировка:\n")
            f.write(traceback.format_exc())

        return None


def create_correctness_summary(logger=None, is_incoming=False):
    """
    Создает таблицу 'Сводная корректности' с формулами.

    Args:
        logger: логгер для записи хода выполнения
        is_incoming: True для входящих, False для исходящих

    Returns:
        Кортеж (DataFrame с данными, список формул Excel, список комментариев).
    """
    if logger:
        logger.info("Создание таблицы 'Сводная корректности'...")

    table_data = []
    formulas = []
    formula_comments = []

    # 1. Диалог отсутствует - ячейка N4
    col_dur = 'K' if is_incoming else 'M'
    excel_formula_no_dialog = f"=СЧЁТЕСЛИ('Исходный отчёт'!{col_dur}:{col_dur},0)"
    comment_no_dialog = "Количество значений в листе 'Исходный отчёт', где 'Голосовой робот длительность' = 0"
    table_data.append({'Корректность': 'Диалог отсутствует', 'Количество': excel_formula_no_dialog})
    formulas.append(excel_formula_no_dialog)
    formula_comments.append(comment_no_dialog)

    # 2. Ошибка - ячейка N5
    col_err = 'T' if is_incoming else 'V'
    excel_formula_error = f'=СЧЁТЕСЛИ(\'Отчёт по ошибкам\'!{col_err}:{col_err},"Ошибка")'
    comment_error = "Количество значений в листе 'Отчёт по ошибкам', где 'Корректность' = 'Ошибка'"
    table_data.append({'Корректность': 'Ошибка', 'Количество': excel_formula_error})
    formulas.append(excel_formula_error)
    formula_comments.append(comment_error)

    # 3. Корректно - ячейка N6
    excel_formula_correct = '=B14-N5'
    comment_correct = "Итого звонков из таблицы 'Результаты в исходном отчёте' - количество ошибок"
    table_data.append({'Корректность': 'Корректно', 'Количество': excel_formula_correct})
    formulas.append(excel_formula_correct)
    formula_comments.append(comment_correct)

    # Создаем DataFrame
    correctness_df = pd.DataFrame(table_data)

    if logger:
        logger.info("✅ Таблица 'Сводная корректности' создана")

    return correctness_df, formulas, formula_comments


def create_results_summary(df: pd.DataFrame, source_name: str = "", logger=None):
    """
    Создаёт сводную таблицу по результатам звонков: количество и процент
    для каждого уникального результата, плюс строка «Итого».

    Args:
        df: DataFrame с данными
        source_name: название источника для логирования
        logger: логгер для записи хода выполнения

    Returns:
        DataFrame с колонками ['Результат', 'Количество', 'Процент'].
        При ошибке — DataFrame с колонкой 'Ошибка'.
    """
    try:
        column_name = 'results-confirmation'  # ЗАМЕНИЛ COLUMN_CONFIG на прямое название

        if column_name not in df.columns:
            if logger:
                logger.error(f"Колонка '{column_name}' не найдена в данных {source_name}")
            return pd.DataFrame({'Ошибка': [f"Колонка '{column_name}' не найдена"]})

        # Считаем количество по каждому результату
        results_counts = df[column_name].value_counts().reset_index()
        results_counts.columns = ['Результат', 'Количество']

        # Добавляем проценты как десятичные дроби
        total = results_counts['Количество'].sum()
        results_counts['Процент'] = (results_counts['Количество'] / total).round(4)

        # Добавляем строку "Итого"
        итого_строка = pd.DataFrame({
            'Результат': ['Итого'],
            'Количество': [total],
            'Процент': [1.0]  # 1.0 = 100%
        })

        results_summary = pd.concat([results_counts, итого_строка], ignore_index=True)

        if logger:
            logger.info(
                f"Создана сводная по результатам ({source_name}): {len(results_summary) - 1} результатов + Итого")
        return results_summary

    except Exception as e:
        if logger:
            logger.error(f"Ошибка при создании сводной по результатам ({source_name}): {e}")
        return pd.DataFrame({'Ошибка': [f'Ошибка обработки: {str(e)[:100]}']})


def create_last_calls_report(df: pd.DataFrame, logger=None, is_incoming=False):
    """
    Создает отчет по последнему звонку для каждого уникального номера.
    Дата ожидается в формате 'YYYY-MM-DD HH:MM:SS'

    Args:
        df: DataFrame с исходными данными
        logger: логгер для записи хода выполнения
        is_incoming: True для входящих, False для исходящих

    Returns:
        DataFrame с последними звонками по каждому номеру.
        Пустой DataFrame, если данных нет.
    """
    try:
        phone_column = 'Вызывающий номер' if is_incoming else 'Вызываемый номер'
        date_column = 'Дата начала звонка'

        if phone_column not in df.columns or date_column not in df.columns:
            if logger:
                logger.warning(f"Не найдены колонки: {phone_column} или {date_column}")
            return pd.DataFrame()

        if logger:
            logger.info("Создание отчёта по последним звонкам...")
            logger.info(f"Колонка номера: '{phone_column}'")
            logger.info(f"Колонка даты: '{date_column}'")

        df_clean = df.copy()

        # Удаляем строки с пустым номером
        df_clean = df_clean[df_clean[phone_column].notna()]

        # Парсим дату ТОЛЬКО в формате ISO (без dayfirst)
        df_clean['_temp_date'] = pd.to_datetime(df_clean[date_column], errors='coerce')

        # Удаляем строки с невалидной датой
        invalid_mask = df_clean['_temp_date'].isna()
        if invalid_mask.any() and logger:
            invalid_examples = df_clean[invalid_mask][date_column].head(3).tolist()
            logger.warning(f"Удалено {invalid_mask.sum()} строк с невалидными датами. Примеры: {invalid_examples}")

        df_clean = df_clean[~invalid_mask].copy()

        if df_clean.empty:
            if logger:
                logger.warning("Нет строк с валидными датами")
            return pd.DataFrame()

        # Сортируем по убыванию даты и оставляем последний звонок для каждого номера
        df_clean = df_clean.sort_values('_temp_date', ascending=False)
        last_calls = df_clean.drop_duplicates(subset=[phone_column], keep='first')

        last_calls = last_calls.drop(columns=['_temp_date'])

        # Сортируем результат по дате (от новых к старым)
        last_calls = last_calls.sort_values(
            date_column,
            key=lambda col: pd.to_datetime(col, errors='coerce'),
            ascending=False
        )

        unique_numbers = last_calls[phone_column].nunique()
        if logger:
            logger.info(f"✅ Отчёт по последним звонкам создан: {unique_numbers} уникальных номеров")

        return last_calls

    except Exception as e:
        if logger:
            logger.error(f"❌ Ошибка при создании отчёта по последним звонкам: {e}")
            import traceback
            logger.error(traceback.format_exc())
        return pd.DataFrame()


def create_telephony_status_pivot(df: pd.DataFrame, logger=None):
    """
    Создаёт сводную таблицу по статусам телефонии с процентами
    и итоговой строкой «Итого».

    Args:
        df: DataFrame с колонкой 'Статус'
        logger: логгер для записи хода выполнения

    Returns:
        DataFrame с колонками ['Статус', 'Количество', 'Процент'].
        При ошибке — DataFrame с колонкой 'Ошибка'.
    """
    try:
        column_name = 'Статус'  # Прямое название вместо COLUMN_CONFIG

        if column_name not in df.columns:
            if logger:
                logger.error(f"Колонка '{column_name}' не найдена в данных")
            return pd.DataFrame({'Ошибка': [f"Колонка '{column_name}' не найдена"]})

        # Считаем количество по каждому статусу
        status_counts = df[column_name].value_counts().reset_index()
        status_counts.columns = ['Статус', 'Количество']

        # Добавляем проценты как десятичные дроби
        total = status_counts['Количество'].sum()
        status_counts['Процент'] = (status_counts['Количество'] / total).round(4)

        # Добавляем строку "Итого"
        итого_строка = pd.DataFrame({
            'Статус': ['Итого'],
            'Количество': [total],
            'Процент': [1.0]
        })

        status_counts = pd.concat([status_counts, итого_строка], ignore_index=True)

        if logger:
            logger.info(f"Создана сводная по статусам: {len(status_counts) - 1} статусов + Итого")
        return status_counts

    except Exception as e:
        if logger:
            logger.error(f"Ошибка при создании сводной по статусам: {e}")
        return pd.DataFrame({'Ошибка': [f'Ошибка обработки: {str(e)[:100]}']})


def create_auto_answer_pivot(df: pd.DataFrame, logger=None):
    """
    Создаёт сводную таблицу по результатам детектора автоответчиков.
    Используется только для исходящих отчётов.

    Args:
        df: DataFrame с колонкой 'Результат AMD'
        logger: логгер для записи хода выполнения

    Returns:
        DataFrame с колонками ['Результат AMD', 'Количество', 'Процент'].
        При ошибке — DataFrame с колонкой 'Ошибка'.
    """
    try:
        column_name = 'Результат AMD'  # Прямое название вместо COLUMN_CONFIG

        if column_name not in df.columns:
            if logger:
                logger.error(f"Колонка '{column_name}' не найдена в данных")
            return pd.DataFrame({'Ошибка': [f"Колонка '{column_name}' не найдена"]})

        # Считаем количество по каждому типу ответа
        auto_counts = df[column_name].value_counts().reset_index()
        auto_counts.columns = ['Результат AMD', 'Количество']

        # Добавляем проценты как десятичные дроби
        total = auto_counts['Количество'].sum()
        auto_counts['Процент'] = (auto_counts['Количество'] / total).round(4)

        # Добавляем строку "Итого"
        итого_строка = pd.DataFrame({
            'Результат AMD': ['Итого'],
            'Количество': [total],
            'Процент': [1.0]
        })

        auto_counts = pd.concat([auto_counts, итого_строка], ignore_index=True)

        if logger:
            logger.info(f"Создана сводная по автоответчику: {len(auto_counts) - 1} типов + Итого")
        return auto_counts

    except Exception as e:
        if logger:
            logger.error(f"Ошибка при создании сводной по автоответчику: {e}")
        return pd.DataFrame({'Ошибка': [f'Ошибка обработки: {str(e)[:100]}']})


def add_repeat_count_column(ws_last_calls, is_incoming, logger=None):
    """
    Добавляет на лист 'Отчёт по последнему звонку' колонку 'Количество перезвонов'
    с формулой СЧЁТСЛИ, считающей количество звонков по номеру в исходном отчёте.
    Изменяет переданный лист на месте, ничего не возвращает.

    Args:
        ws_last_calls: лист 'Отчёт по последнему звонку' (openpyxl worksheet)
        is_incoming: True для входящих, False для исходящих
        logger: опциональный логгер

    Returns:
        None
    """
    # Определяем название колонки с номером в исходном отчёте
    if is_incoming:
        phone_col_name = 'Вызывающий номер'
    else:
        phone_col_name = 'Вызываемый номер'

    header_row = 4  # начиная с 4 строки заголовки колонок
    col_letter = None
    for col in range(1, ws_last_calls.max_column + 1):
        cell_value = ws_last_calls.cell(row=header_row, column=col).value
        if cell_value == phone_col_name:
            col_letter = get_column_letter(col)
            break

    if col_letter is None:
        if logger:
            logger.warning(f"Не найдена колонка '{phone_col_name}' на листе 'Отчёт по последнему звонку'")
        return

    if is_incoming:
        original_col_letter = 'F'  # Вызывающий номер в исходном отчёте
    else:
        original_col_letter = 'D'  # Вызываемый номер в исходном отчёте

    target_col_num = None
    for col in range(1, ws_last_calls.max_column + 1):
        if ws_last_calls.cell(row=header_row, column=col).value == phone_col_name:
            target_col_num = col
            break

    if target_col_num is None:
        return

    # Вставляем новую колонку после target_col_num
    insert_col_num = target_col_num + 1
    ws_last_calls.insert_cols(insert_col_num)

    # Заголовок новой колонки
    ws_last_calls.cell(row=header_row, column=insert_col_num, value="Количество перезвонов")

    # Формула для каждой строки данных (начиная с 5-й строки, т.к. строки 1-4 служебные)
    data_start_row = header_row + 1  # 5
    max_row = ws_last_calls.max_row
    for row in range(data_start_row, max_row + 1):
        phone_cell = ws_last_calls.cell(row=row, column=target_col_num)
        formula = f"=СЧЁТЕСЛИ('Исходный отчёт'!{original_col_letter}:{original_col_letter}, {phone_cell.coordinate})"
        ws_last_calls.cell(row=row, column=insert_col_num, value=formula)

    if logger:
        logger.info(f"✅ Добавлена колонка 'Количество перезвонов' (столбец {get_column_letter(insert_col_num)})")


def format_excel_table(ws, start_row: int, start_col: int = 1, is_results_table: bool = False, logger=None):
    """
    Применяет форматирование к таблице Excel: заголовки, границы,
    автоширину колонок, подсветку успешных результатов и строки «Итого».

    Args:
        ws: лист Excel (openpyxl worksheet)
        start_row: номер строки, с которой начинается таблица (строка заголовков)
        start_col: номер первой колонки таблицы (по умолчанию 1)
        is_results_table: True, если таблица содержит результаты звонков
                          (тогда успешные строки подсвечиваются зелёным)
        logger: логгер для записи предупреждений

    Returns:
        None
    """
    try:
        # Определяем размеры таблицы
        max_row = ws.max_row
        max_col = ws.max_column

        # Стили
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        header_alignment = Alignment(horizontal="center", vertical="center")

        data_font = Font(name="Calibri", size=11)
        data_alignment = Alignment(vertical="center")
        number_alignment = Alignment(horizontal="right", vertical="center")

        # Зеленый фон для успешных результатов - ОБНОВЛЁННЫЙ СПИСОК
        success_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        success_font = Font(name="Calibri", size=11, bold=True)

        border_side = Side(border_style="thin", color="000000")
        border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)

        # Форматируем заголовки
        for col in range(start_col, max_col + 1):
            cell = ws.cell(row=start_row, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = border

            # Автоподбор ширины колонки
            column_letter = get_column_letter(col)
            max_length = 0

            for row in range(start_row, max_row + 1):
                cell_value = ws.cell(row=row, column=col).value
                if cell_value:
                    length = len(str(cell_value))
                    if length > max_length:
                        max_length = length

            # Устанавливаем ширину с небольшим запасом
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width

        # Форматируем данные
        for row in range(start_row + 1, max_row + 1):
            is_success_row = False

            # Проверяем, является ли эта строка успешным результатом
            if is_results_table:
                result_cell = ws.cell(row=row, column=start_col)
                result_value = str(result_cell.value) if result_cell.value else ""

                # Проверяем вхождение в ОБНОВЛЁННЫЙ список успешных результатов
                if any(success_result == result_value for success_result in SUCCESSFUL_RESULTS):
                    is_success_row = True

            for col in range(start_col, max_col + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = border

                # Применяем стиль в зависимости от результата
                if is_success_row:
                    cell.fill = success_fill
                    cell.font = success_font
                else:
                    cell.font = data_font

                # Определяем тип данных для выравнивания
                cell_value = cell.value

                # ЕСЛИ ЭТО КОЛОНКА "Процент" - ФОРМАТ ПРОЦЕНТОВ
                header_cell = ws.cell(row=start_row, column=col)
                if header_cell.value == 'Процент':
                    cell.number_format = '0.00%'
                    cell.alignment = number_alignment
                elif isinstance(cell_value, (int, float)):
                    cell.alignment = number_alignment
                else:
                    cell.alignment = data_alignment

        # Выделяем строку "Итого" если она есть
        for row in range(start_row + 1, max_row + 1):
            cell = ws.cell(row=row, column=start_col)
            if cell.value == "Итого":
                # Форматируем всю строку "Итого"
                for col in range(start_col, max_col + 1):
                    total_cell = ws.cell(row=row, column=col)
                    total_cell.font = Font(bold=True)
                    total_cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                break

    except Exception as e:
        if logger:
            logger.warning(f"Не удалось применить форматирование: {e}")
        else:
            logging.getLogger(__name__).warning(f"Не удалось применить форматирование: {e}")


def save_to_excel_with_tables_pivot_only(
        df: pd.DataFrame,
        last_calls_df: pd.DataFrame,
        original_results_summary: pd.DataFrame,
        last_calls_results_summary: pd.DataFrame,
        output_path: str,
        logger=None,
        is_incoming=False
) -> None:
    """
    Сохраняет отчёт в режиме «Только сводные»: 3 таблицы на листе «Сводная»
    плюс листы «Сводная телефонии», «Отчёт по последнему звонку» и «Исходный отчёт».

    Args:
        df: исходный DataFrame с данными отчёта
        last_calls_df: DataFrame с последними звонками по каждому номеру
        original_results_summary: сводная по результатам в исходном отчёте
        last_calls_results_summary: сводная по результатам в отчёте по последнему звонку
        output_path: путь для сохранения итогового файла
        logger: логгер для записи хода выполнения
        is_incoming: True для входящих звонков, False для исходящих

    Returns:
        None
    """
    if logger:
        logger.info(f"Сохранение отчёта (только сводные) в Excel: {output_path}")

    # Создаем новую книгу Excel
    wb = Workbook()

    # Удаляем дефолтный лист
    if 'Sheet' in wb.sheetnames:
        default_sheet = wb['Sheet']
        wb.remove(default_sheet)

    # ===== 1. ЛИСТ "Сводная" =====
    ws_summary = wb.create_sheet(title='Сводная')

    # ТАБЛИЦА 1: Результаты в исходном отчёте
    row = 1
    start_col = 1
    ws_summary.cell(row=row, column=start_col, value='Результаты в исходном отчёте').font = Font(bold=True, size=12)
    row += 2

    if not original_results_summary.empty and 'Ошибка' not in original_results_summary.columns:
        headers = list(original_results_summary.columns)
        for col_idx, header in enumerate(headers, 0):
            ws_summary.cell(row=row, column=start_col + col_idx, value=header)

        table1_start_row = row
        row += 1

        for _, data_row in original_results_summary.iterrows():
            for col_idx, value in enumerate(data_row.tolist(), 0):
                ws_summary.cell(row=row, column=start_col + col_idx, value=value)
            row += 1

        format_excel_table(ws_summary, table1_start_row, start_col=start_col, is_results_table=True, logger=logger)

    # ПУСТОЙ СТОЛБЕЦ D

    # ТАБЛИЦА 2: Результаты в отчёте по последнему звонку
    row = 1
    start_col = 5
    ws_summary.cell(row=row, column=start_col, value='Результаты в отчёте по последнему звонку').font = Font(bold=True,
                                                                                                             size=12)
    row += 2

    if not last_calls_results_summary.empty and 'Ошибка' not in last_calls_results_summary.columns:
        headers = list(last_calls_results_summary.columns)
        for col_idx, header in enumerate(headers, 0):
            ws_summary.cell(row=row, column=start_col + col_idx, value=header)

        table2_start_row = row
        row += 1

        for _, data_row in last_calls_results_summary.iterrows():
            for col_idx, value in enumerate(data_row.tolist(), 0):
                ws_summary.cell(row=row, column=start_col + col_idx, value=value)
            row += 1

        format_excel_table(ws_summary, table2_start_row, start_col=start_col, is_results_table=True, logger=logger)

    # ПУСТОЙ СТОЛБЕЦ H

    # ТАБЛИЦА 3: Общая
    # Создаем таблицу "Общая"
    general_df, general_formulas = create_general_table(
        df,
        last_calls_df,
        logger if logger else logging.getLogger(__name__),
        include_correct_calls=False,
        is_incoming=is_incoming
    )

    # Обновляем значения контактности
    if not general_df.empty:
        contact_calls_idx = general_df[general_df['Показатель'] == 'Контактность относительно всех звонков'].index
        contact_numbers_idx = general_df[general_df['Показатель'] == 'Контактность относительно всех номеров'].index

        if len(contact_calls_idx) > 0:
            contact_calls = calculate_contactness(original_results_summary)
            general_df.at[contact_calls_idx[0], 'Значение'] = contact_calls
            calls_count = general_df.at[1, 'Значение']
            if calls_count > 0:
                general_df.at[contact_calls_idx[0], 'Процент'] = contact_calls / calls_count

        if len(contact_numbers_idx) > 0:
            contact_numbers = calculate_contactness(last_calls_results_summary)
            general_df.at[contact_numbers_idx[0], 'Значение'] = contact_numbers
            candidates_count = general_df.at[0, 'Значение']
            if candidates_count > 0:
                general_df.at[contact_numbers_idx[0], 'Процент'] = contact_numbers / candidates_count

    row = 1
    start_col = 9
    ws_summary.cell(row=row, column=start_col, value='Общая').font = Font(bold=True, size=12)
    row += 2

    if not general_df.empty:
        headers = ['Показатель', 'Значение', 'Процент']
        for col_idx, header in enumerate(headers, 0):
            ws_summary.cell(row=row, column=start_col + col_idx, value=header)

        table3_start_row = row
        row += 1

        for idx, data_row in general_df.iterrows():
            # Показатель
            ws_summary.cell(row=row, column=start_col, value=data_row['Показатель'])

            # ЗНАЧЕНИЕ - сюда пишем либо формулу, либо число
            value = data_row['Значение']
            value_cell = ws_summary.cell(row=row, column=start_col + 1)

            if isinstance(value, str) and value.startswith('='):
                value_cell.value = value  # формула Excel
            else:
                value_cell.value = value  # обычное число

            if idx < len(general_formulas) and general_formulas[idx]:
                add_comment_to_cell(value_cell, general_formulas[idx], logger=logger)

            # ПРОЦЕНТ - только для показателей, где он нужен
            percent_cell = ws_summary.cell(row=row, column=start_col + 2)
            indicator = data_row['Показатель']

            if indicator == 'Контактность относительно всех звонков':
                percent_cell.value = '=J7/J5'
                percent_cell.number_format = '0.00%'
                add_comment_to_cell(percent_cell, "Контактность относительно всех звонков / Количество звонков",
                                    logger=logger)

            elif indicator == 'Контактность относительно всех номеров':
                percent_cell.value = '=J8/J4'
                percent_cell.number_format = '0.00%'
                add_comment_to_cell(percent_cell, "Контактность относительно всех номеров / Количество кандидатов",
                                    logger=logger)

            elif indicator == 'Количество корректных звонков':
                percent_cell.value = '=1 - N5 / B14'
                percent_cell.number_format = '0.00%'
                add_comment_to_cell(
                    percent_cell,
                    "1 - Количество ошибок / Количество звонков с каким-либо результатом",
                    logger=logger,
                )

            else:
                # Для всех остальных строк (CPL, Всего лидов, Количество ошибочных лидов и т.д.)
                # процент не нужен - оставляем пустым
                if pd.notna(data_row['Процент']) and data_row['Процент'] != '':
                    percent_cell.value = data_row['Процент']
                    if isinstance(data_row['Процент'], (int, float)):
                        percent_cell.number_format = '0.00%'

            row += 1

        format_excel_table(ws_summary, table3_start_row, start_col=start_col, is_results_table=False, logger=logger)

    if logger:
        logger.info(f"✅ Создан лист 'Сводная' с 3 таблицами")

    # ===== 2. ЛИСТ "Сводная телефонии" =====
    ws_telephony = wb.create_sheet('Сводная телефонии')
    status_pivot = create_telephony_status_pivot(df, logger)

    # Таблица 1: Сводная по статусу телефонии – всегда
    row = 1
    ws_telephony.cell(row, 1, 'Сводная по статусу телефонии').font = Font(bold=True, size=12)
    row += 2
    if not status_pivot.empty and 'Ошибка' not in status_pivot.columns:
        headers = list(status_pivot.columns)
        for col_idx, header in enumerate(headers, 1):
            ws_telephony.cell(row, col_idx, header)
        start_row = row
        row += 1
        for _, data_row in status_pivot.iterrows():
            for col_idx, value in enumerate(data_row.tolist(), 1):
                ws_telephony.cell(row, col_idx, value)
            row += 1
        format_excel_table(ws_telephony, start_row, is_results_table=False, logger=logger)

    # Таблица 2: Сводная по детектору автоответчиков – только для исходящих
    if not is_incoming:
        auto_pivot = create_auto_answer_pivot(df, logger)
        row += 3
        ws_telephony.cell(row, 1, 'Сводная по детектору автоответчиков').font = Font(bold=True, size=12)
        row += 2
        if not auto_pivot.empty and 'Ошибка' not in auto_pivot.columns:
            headers = list(auto_pivot.columns)
            for col_idx, header in enumerate(headers, 1):
                ws_telephony.cell(row, col_idx, header)
            start_row = row
            row += 1
            for _, data_row in auto_pivot.iterrows():
                for col_idx, value in enumerate(data_row.tolist(), 1):
                    ws_telephony.cell(row, col_idx, value)
                row += 1
            format_excel_table(ws_telephony, start_row, is_results_table=False, logger=logger)

    if logger:
        logger.info("✅ Создан лист 'Сводная телефонии'")

    # ===== 3. ЛИСТ "Отчёт по последнему звонку" =====
    if not last_calls_df.empty:
        ws_last_calls = wb.create_sheet(title='Отчёт по последнему звонку')

        ws_last_calls.cell(row=1, column=1, value='Отчёт по последнему звонку').font = Font(bold=True, size=14)
        ws_last_calls.cell(row=2, column=1, value=f'Всего уникальных номеров: {len(last_calls_df)}')

        row = 4
        headers = list(last_calls_df.columns)
        for col_idx, header in enumerate(headers, 1):
            ws_last_calls.cell(row=row, column=col_idx, value=header)

        start_row = row
        row += 1

        for _, data_row in last_calls_df.iterrows():
            for col_idx, value in enumerate(data_row.tolist(), 1):
                ws_last_calls.cell(row=row, column=col_idx, value=value)
            row += 1

        add_repeat_count_column(ws_last_calls, is_incoming, logger)

        format_excel_table(ws_last_calls, start_row, logger=logger)

        if logger:
            logger.info(f"✅ Создан лист 'Отчёт по последнему звонку' ({len(last_calls_df)} строк)")
    else:
        if logger:
            logger.info("⚠️ Лист 'Отчёт по последнему звонку' не создан (нет данных)")

    # ===== 4. ЛИСТ "Исходный отчёт" =====
    ws_original = wb.create_sheet(title='Исходный отчёт')

    headers = list(df.columns)
    for col_idx, header in enumerate(headers, 1):
        ws_original.cell(row=1, column=col_idx, value=header)

    for row_idx, (_, data_row) in enumerate(df.iterrows(), 2):
        for col_idx, value in enumerate(data_row.tolist(), 1):
            ws_original.cell(row=row_idx, column=col_idx, value=value)

    format_excel_table(ws_original, 1, logger=logger)

    if logger:
        logger.info("✅ Создан лист 'Исходный отчёт'")

    # Сохраняем файл
    wb.save(output_path)
    if logger:
        logger.info(f"✅ Файл сохранен: {output_path}")
        logger.info(f"   Листы в файле: {wb.sheetnames}")


def calculate_contactness(results_summary_df):
    """
    Вычисляет контактность как разницу между общим количеством
    и суммой неконтактных результатов (автоответчики, тишина, сбросы).

    Args:
        results_summary_df: сводная таблица с колонками 'Результат' и 'Количество'

    Returns:
        Число контактных звонков (int). 0, если данных нет или нет строки «Итого».
    """
    if results_summary_df.empty or 'Количество' not in results_summary_df.columns:
        return 0

    # Находим строку "Итого"
    total_row = results_summary_df[results_summary_df['Результат'] == 'Итого']
    if total_row.empty:
        return 0

    total = total_row['Количество'].iloc[0]

    # Находим количество неконтактных результатов - ИСПОЛЬЗУЕМ NON_CONTACT_RESULTS
    # Преобразуем к нижнему регистру для сравнения без учёта регистра
    non_contact_mask = results_summary_df['Результат'].apply(
        lambda x: str(x).strip().lower() if pd.notna(x) else ''
    ).isin([r.lower() for r in NON_CONTACT_RESULTS])

    non_contact_rows = results_summary_df[non_contact_mask]
    non_contact_count = non_contact_rows['Количество'].sum() if not non_contact_rows.empty else 0

    # Контактность = Итого - Неконтактные
    return total - non_contact_count


def check_columns(input_excel_path: str) -> Dict[str, bool]:
    """
    Проверяет наличие обязательных колонок в Excel-файле перед обработкой.

    Args:
        input_excel_path: путь к Excel-файлу

    Returns:
        Словарь {название_колонки: True/False} для каждой проверяемой колонки.
        При ошибке чтения файла возвращает {'error': 'описание'}.
    """
    try:
        # Читаем только заголовки
        df = pd.read_excel(input_excel_path, nrows=0)

        # Прямые названия колонок вместо COLUMN_CONFIG
        required_columns = [
            'Статус',  # для сводной по статусам
            'results-confirmation',  # для сводных по результатам
            'Вызываемый номер',  # для отчёта по последнему звонку
            'Дата начала звонка'  # для отчёта по последнему звонку
        ]

        results = {}
        for column_name in required_columns:
            results[column_name] = column_name in df.columns

        return results

    except Exception as e:
        return {'error': str(e)}


def generate_pivot_report(input_excel_path: str, logger=None) -> Tuple[str, str]:
    """
    Основная функция режима «Только сводные»: загружает отчёт, строит
    сводные таблицы по статусам, автоответчику, результатам и последним
    звонкам, сохраняет результат в Excel.

    Args:
        input_excel_path: путь к исходному Excel-файлу
        logger: опциональный логгер. Если не передан — создаётся автоматически.

    Returns:
        Кортеж (путь_к_отчёту, путь_к_лог_файлу) при успехе,
        (None, путь_к_файлу_ошибки) при ошибке.
    """
    # Если логгер не передан, создаем свой
    if logger is None:
        logger = logging.getLogger(f"pivot_{datetime.now().strftime('%d.%m.%Y_%H-%M')}")
        logger.setLevel(logging.INFO)

    # СОЗДАЁМ ЛОГГЕР для файла
    timestamp = datetime.now().strftime('%d.%m.%Y_%H-%M')
    original_name = Path(input_excel_path).stem
    file_prefix = f"pivot_{timestamp}_{original_name}"
    log_filename = f"logs/{file_prefix}.log"
    error_filename = f"logs/ERROR_{file_prefix}.txt"

    Path("logs").mkdir(exist_ok=True)

    # Настраиваем логгер для записи в файл
    file_handler = logging.FileHandler(log_filename, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    logger.info("=" * 60)
    logger.info(f"НАЧАЛО СОЗДАНИЯ СВОДНЫХ ТАБЛИЦ: {timestamp}")
    logger.info(f"Исходный файл: {input_excel_path}")
    logger.info(f"Лог-файл: {log_filename}")
    logger.info("=" * 60)

    try:
        # ЗАГРУЖАЕМ EXCEL-ФАЙЛ
        logger.info(f"Загрузка файла: {input_excel_path}")
        df = pd.read_excel(input_excel_path)
        logger.info(f"✅ Загружено {len(df)} строк, {len(df.columns)} колонок")

        df = filter_test_phones(df, logger)  # УДАЛЯЕМ ТЕСТОВЫЕ НОМЕРА

        # Удаляем ненужные колонки
        columns_to_remove = ['Идентификатор задания', 'Имя задания']
        existing_columns = [col for col in columns_to_remove if col in df.columns]

        if existing_columns:
            df = df.drop(columns=existing_columns)
            logger.info(f"✅ Удалены колонки: {', '.join(existing_columns)}")
            logger.info(f"   Осталось колонок: {len(df.columns)}")

        # ПРОВЕРКА НЕОБХОДИМЫХ КОЛОНОК
        required_cols = [
            'Статус',
            'results-confirmation',
            'Вызываемый номер',
            'Дата начала звонка'
        ]

        missing_cols = [col for col in required_cols if col not in df.columns]

        if missing_cols:
            error_msg = f"❌ Отсутствуют необходимые колонки: {missing_cols}"
            logger.error(error_msg)
            logger.info(f"Доступные колонки: {list(df.columns)}")
            raise ValueError(error_msg)

        logger.info("✅ Все необходимые колонки присутствуют")

        # СОЗДАЁМ СВОДНЫЕ ТАБЛИЦЫ
        logger.info("Создание сводных таблиц...")
        start_time = datetime.now()

        # Отчет по последнему звонку
        is_incoming = is_incoming_report(df)
        last_calls_df = create_last_calls_report(df, logger, is_incoming=is_incoming)

        # Сводная по результатам в исходном отчете
        original_results_summary = create_results_summary(df, "Исходный отчёт", logger)

        # Сводная по результатам в отчете по последнему звонку
        last_calls_results_summary = create_results_summary(last_calls_df, "Отчёт по последнему звонку", logger)

        # СОХРАНЕНИЕ РЕЗУЛЬТАТА
        processed_dir = "processed"
        Path(processed_dir).mkdir(exist_ok=True)
        output_filename = f"{processed_dir}/{file_prefix}_processed.xlsx"

        logger.info(f"Сохранение результата в: {output_filename}")

        save_to_excel_with_tables_pivot_only(
            df,
            last_calls_df,
            original_results_summary,
            last_calls_results_summary,
            output_filename,
            logger,
            is_incoming=is_incoming
        )

        # СТАТИСТИКА
        total_time = (datetime.now() - start_time).total_seconds()
        logger.info("=" * 60)
        logger.info("📊 ИТОГИ СОЗДАНИЯ СВОДНЫХ ТАБЛИЦ:")
        logger.info(f"   Время обработки: {total_time:.1f} секунд")
        logger.info(f"   Файл с результатом: {output_filename}")
        logger.info(f"   Лог-файл: {log_filename}")
        logger.info("=" * 60)
        logger.info("✅ Сводные таблицы созданы успешно!")

        # Убираем хендлер чтобы не дублировалось
        logger.removeHandler(file_handler)

        return output_filename, log_filename

    except Exception as e:
        # Если произошла ошибка - логируем
        logger.error("❌ КРИТИЧЕСКАЯ ОШИБКА:")
        logger.error(f"   Тип ошибки: {type(e).__name__}")
        logger.error(f"   Сообщение: {str(e)}")
        logger.error("=" * 60)

        # Создаём файл с подробной информацией об ошибке
        with open(error_filename, 'w', encoding='utf-8') as f:
            f.write(f"Ошибка при создании сводных таблиц {input_excel_path}:\n")
            f.write(f"Время: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
            f.write(f"Тип ошибки: {type(e).__name__}\n")
            f.write(f"Сообщение: {str(e)}\n\n")
            import traceback
            f.write("Трассировка:\n")
            f.write(traceback.format_exc())

        # Убираем хендлер
        logger.removeHandler(file_handler)

        return None, error_filename
