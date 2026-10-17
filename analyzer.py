import pandas as pd
import openai
import json
import time
import logging
import traceback
import re
from pathlib import Path
from datetime import datetime

from config import Config
from pivot_generator import filter_test_phones


# ===== КЛАСС ДЛЯ ОШИБКИ ИСТЕЧЕНИЯ API-КЛЮЧА =====
class APIKeyExpiredError(Exception):
    """Исключение для случая истечения срока действия API-ключа"""
    pass


# ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =====
def clean_json_response(text):
    """Очищает ответ агента от markdown кавычек"""
    if not text:
        return text

    cleaned = text.strip()
    if cleaned.startswith('```json'):
        cleaned = cleaned[7:]
    elif cleaned.startswith('```'):
        cleaned = cleaned[3:]

    if cleaned.endswith('```'):
        cleaned = cleaned[:-3]

    return cleaned.strip()


def create_batch_input(batch_data):
    """Создает входные данные для пакета диалогов"""
    input_text = ''
    for i, data in enumerate(batch_data, 1):
        input_text += f"ДИАЛОГ {i}\n"
        input_text += f"Текст разговора: {data['dialogue']}\n"
        input_text += f"Результат диалога (Twin): {data['result']}\n\n"
    return input_text


def filter_report(df, logger, column_name='results-confirmation'):
    """
    Фильтрует DataFrame, удаляя строки с определёнными значениями
    в указанной колонке (неинтересные диалоги), а также пустые
    и NaN-значения.

    Args:
        df: DataFrame для фильтрации
        logger: логгер для записи информации
        column_name: название колонки для фильтрации

    Returns:
        Отфильтрованный DataFrame с обнулённым индексом.
    """
    # Значения для удаления (регистр важен!)
    values_to_remove = [
        'Автоответчик',
        'Тишина',
        'Сброс',
        'Поступил звонок',
        'Спам',
        'Умный автоответчик',
        'Попали в компанию',
        'Лид в браке'
    ]

    # Проверяем наличие колонки
    if column_name not in df.columns:
        logger.warning(f"Колонка '{column_name}' не найдена, фильтрация не применена")
        return df

    # Запоминаем исходное количество строк
    original_count = len(df)

    # Копируем DataFrame чтобы не менять оригинальный
    df_filtered = df.copy()

    # Удаляем строки:
    # 1. С определенными значениями (точное совпадение с регистром)
    # 2. С пустыми значениями (NaN или пустая строка)

    condition = df_filtered[column_name].apply(lambda x:
                                               (isinstance(x, str) and x.strip() in values_to_remove) or
                                               (pd.isna(x)) or
                                               (isinstance(x, str) and x.strip() == '')
                                               )

    # Фильтруем (оставляем только НЕ подходящие под условие удаления)
    df_filtered = df_filtered[~condition].copy()

    # Сбрасываем индекс
    df_filtered = df_filtered.reset_index(drop=True)

    # Логируем результат фильтрации
    removed_count = original_count - len(df_filtered)
    logger.info(f"✅ Фильтрация отчёта:")
    logger.info(f"   Исходных строк: {original_count}")
    logger.info(f"   Удалено строк: {removed_count}")
    logger.info(f"   Осталось строк: {len(df_filtered)}")

    if removed_count > 0:
        # Выводим статистику по удаленным значениям
        removed_stats = df[condition][column_name].value_counts()
        logger.info(f"   Статистика удаления:")
        for value, count in removed_stats.items():
            value_str = f"'{value}'" if pd.notna(value) else "[пустое/NaN]"
            logger.info(f"     {value_str}: {count} строк")

    return df_filtered


def check_api_key_error(error_msg):
    """Проверяет, является ли ошибка ошибкой истечения API-ключа"""
    if not error_msg:
        return False

    error_msg_lower = str(error_msg).lower()

    # Проверяем различные варианты сообщений об истечении ключа
    key_expired_phrases = [
        'apikey has expired',
        'api key has expired',
        'api_key has expired',
        'key has expired',
        'expired apikey',
        'expired api key',
        'token has expired',
        'authentication failed',
        'unauthorized',
        'unaauthenticated',
        'invalid api key',
        'invalid api-key',
        'invalid apikey',
        '401'
    ]

    for phrase in key_expired_phrases:
        if phrase in error_msg_lower:
            return True
    return False


def analyze_batch(batch_data, logger, client, prompt_id):
    """
    Анализирует пакет диалогов через нейросеть Yandex Cloud.
    Отправляет диалоги в LLM, парсит ответ, возвращает результаты по каждому.

    Args:
        batch_data: список словарей с данными диалогов
                    (excel_row, df_index, call_id, dialogue, result)
        logger: логгер для записи хода выполнения
        client: клиент OpenAI (настроенный на Yandex Cloud)
        prompt_id: ID промпта-агента в AI Studio

    Returns:
        Список словарей с результатами: excel_row, df_index,
        correctness, error_type.

    Raises:
        APIKeyExpiredError: если API-ключ истёк или недействителен.
    """
    try:
        logger.info(f"Отправка пакета из {len(batch_data)} диалогов...")
        start_batch_time = time.time()

        user_input = create_batch_input(batch_data)
        response = client.responses.create(
            prompt={"id": prompt_id},
            input=user_input
        )

        batch_time = time.time() - start_batch_time
        logger.info(f"Время обработки пачки: {batch_time:.2f} сек")

        if not hasattr(response, 'output_text') or not response.output_text:
            raise ValueError("Пустой ответ от агента")

        response_text = response.output_text

        # ВАЖНО: проверяем сам ответ от API на ошибку истечения ключа
        if check_api_key_error(response_text):
            logger.critical(f"⛔️ В ОТВЕТЕ АГЕНТА ОБНАРУЖЕНА ОШИБКА ИСТЕЧЕНИЯ API-КЛЮЧА!")
            logger.critical(f"Ответ агента: {response_text[:500]}...")
            raise APIKeyExpiredError(f"API-ключ истек (в ответе агента): {response_text[:200]}...")

        cleaned_text = clean_json_response(response_text)

        # Проверяем, не является ли cleaned_text сообщением об ошибке
        if check_api_key_error(cleaned_text):
            logger.critical(f"⛔️ В ОЧИЩЕННОМ ОТВЕТЕ ОБНАРУЖЕНА ОШИБКА ИСТЕЧЕНИЯ API-КЛЮЧА!")
            raise APIKeyExpiredError(f"API-ключ истек (в очищенном ответе): {cleaned_text[:200]}...")

        # ПАРСИНГ ОТВЕТА ОТ НЕЙРОСЕТИ
        logger.info(f"Длина ответа от нейросети: {len(cleaned_text)} символов")
        logger.info(f"Первые 300 символов ответа: {cleaned_text[:300]}...")

        try:
            parsed = json.loads(cleaned_text)
            logger.info(f"Тип ответа от нейросети после парсинга JSON: {type(parsed)}")

            if isinstance(parsed, list):
                # Идеальный случай - массив объектов
                results = parsed
                logger.info(f"✅ Нейросеть вернула массив из {len(results)} элементов")

            elif isinstance(parsed, dict):
                # Нейросеть вернула словарь - логируем структуру
                logger.warning(f"⚠️ Нейросеть вернула словарь вместо массива. Ключи: {list(parsed.keys())}")
                logger.info(f"Содержимое словаря: {str(parsed)[:300]}...")

                # Проверяем разные возможные форматы
                if "results" in parsed and isinstance(parsed["results"], list):
                    # Формат: {"results": [...]}
                    results = parsed["results"]
                    logger.info(f"Нашли результаты в ключе 'results': {len(results)} элементов")

                elif "Корректность" in parsed or "Тип_ошибки" in parsed:
                    # Это один результат, оборачиваем в массив
                    results = [parsed]
                    logger.info("Преобразовали единичный результат в массив")

                elif len(parsed) == 1 and isinstance(list(parsed.values())[0], list):
                    # Формат: {"some_key": [...]}
                    key = list(parsed.keys())[0]
                    results = parsed[key]
                    logger.info(f"Нашли результаты в ключе '{key}': {len(results)} элементов")

                else:
                    # Пытаемся найти любой список в значениях
                    found_list = False
                    for key, value in parsed.items():
                        if isinstance(value, list):
                            results = value
                            logger.info(f"Нашли результаты в ключе '{key}': {len(results)} элементов")
                            found_list = True
                            break

                    if not found_list:
                        # Если ничего не нашли - создаем ошибку формата
                        error_msg = f"Нейросеть вернула непонятный формат. Ожидался массив, получен словарь с ключами: {list(parsed.keys())}"
                        logger.error(error_msg)
                        logger.error(f"Полный ответ: {cleaned_text[:1000]}")
                        raise ValueError(error_msg)

            else:
                # Совсем непонятный формат
                error_msg = f"Ответ нейросети не является ни массивом, ни словарем: {type(parsed)}"
                logger.error(error_msg)
                raise ValueError(error_msg)

        except json.JSONDecodeError as e:
            # Детальный лог ошибки парсинга
            logger.error(f"❌ Ошибка парсинга JSON: {e}")
            logger.error(f"Ответ от агента (первые 500 символов): {response_text[:500]}")
            logger.error(f"Очищенный текст (первые 500 символов): {cleaned_text[:500]}")

            # Пробуем найти JSON в тексте
            json_match = re.search(r'\[.*\]', cleaned_text, re.DOTALL)
            if json_match:
                logger.info("Пробую извлечь JSON с помощью регулярного выражения...")
                try:
                    extracted_json = json_match.group(0)
                    parsed = json.loads(extracted_json)
                    if isinstance(parsed, list):
                        results = parsed
                        logger.info(f"✅ Удалось извлечь JSON: {len(results)} элементов")
                    else:
                        raise ValueError(f"Извлеченный JSON не является массивом: {type(parsed)}")
                except Exception as extract_error:
                    logger.error(f"Не удалось извлечь JSON: {extract_error}")
                    raise e
            else:
                raise e

        # ПРОВЕРКА КОЛИЧЕСТВА РЕЗУЛЬТАТОВ
        if not isinstance(results, list):
            raise ValueError(f"Ответ не является массивом: {type(results)}")

        if len(results) != len(batch_data):
            logger.warning(f"⚠️ Агент вернул {len(results)} результатов, ожидалось {len(batch_data)}")

            # Если результатов меньше - дополняем ошибками
            if len(results) < len(batch_data):
                logger.warning(f"Дополняю недостающие {len(batch_data) - len(results)} результатов ошибками")
                for i in range(len(results), len(batch_data)):
                    results.append({
                        "Корректность": "Ошибка",
                        "Тип_ошибки": "Неполный ответ от нейросети"
                    })

        batch_results = []
        for i, (result, call_data) in enumerate(zip(results, batch_data)):
            # Получаем значения из ответа нейросети
            if isinstance(result, dict):
                correctness = result.get("Корректность", "").strip()
                error_type = result.get("Тип_ошибки", "").strip()
            else:
                correctness = "Ошибка"
                error_type = f"Неверный формат элемента: {type(result)}"
                logger.warning(f"Элемент {i} не является словарем: {type(result)}")

            logger.info(
                f"  Диалог {i + 1} (ID: {call_data['call_id']}): Корректность='{correctness}', Тип_ошибки='{error_type}'")

            if correctness == "Корректно":
                error_type = ""
            elif correctness == "Ошибка распознавания":
                error_type = ""
            elif correctness == "Ошибка":
                pass
            else:
                # Любое другое значение — считаем ошибкой нейросети
                correctness = "Ошибка нейросети"
                error_type = "Неверный формат ответа"

            batch_results.append({
                "excel_row": call_data["excel_row"],
                "df_index": call_data["df_index"],
                "correctness": correctness,
                "error_type": error_type
            })

            logger.info(f"  Диалог {i + 1} (ID: {call_data['call_id']}): {correctness}")
            if error_type:
                logger.info(f"    Тип ошибки: {error_type[:80]}...")

        return batch_results

    except openai.APIError as e:
        error_msg = str(e)
        logger.error(f"Ошибка API OpenAI: {error_msg}")

        # Проверяем, истек ли API-ключ
        if check_api_key_error(error_msg):
            logger.critical("⛔️ ОБНАРУЖЕНА КРИТИЧЕСКАЯ ОШИБКА: Истек срок действия API-ключа!")
            raise APIKeyExpiredError(f"API-ключ истек (APIError): {error_msg[:200]}...")

        # Для других API ошибок продолжаем работу, но логируем
        return [
            {
                "excel_row": data["excel_row"],
                "df_index": data["df_index"],
                "correctness": "Ошибка API",
                "error_type": f"Ошибка API: {error_msg[:100]}"
            }
            for data in batch_data
        ]

    except openai.AuthenticationError as e:
        # Ошибка аутентификации - точно истек ключ
        error_msg = str(e)
        logger.critical(f"⛔️ ОШИБКА АУТЕНТИФИКАЦИИ: {error_msg}")
        raise APIKeyExpiredError(f"API-ключ недействителен (AuthenticationError): {error_msg[:200]}...")

    except json.JSONDecodeError as e:
        logger.error(f"Ошибка парсинга JSON: {e}")

        # Проверяем, не содержит ли текст JSON ошибки об истечении ключа
        if check_api_key_error(str(e)) or (response_text and check_api_key_error(response_text)):
            logger.critical("⛔️ ОБНАРУЖЕНА ОШИБКА ИСТЕЧЕНИЯ КЛЮЧА ПРИ ДЕКОДИРОВАНИИ JSON!")
            raise APIKeyExpiredError(f"API-ключ истек (в ответе JSON): {str(e)[:200]}...")

        return [
            {
                "excel_row": data["excel_row"],
                "df_index": data["df_index"],
                "correctness": "Ошибка API",
                "error_type": f"Невалидный JSON от агента: {str(e)[:100]}"
            }
            for data in batch_data
        ]

    except Exception as e:
        error_msg = str(e)
        logger.error(f"Общая ошибка при анализе пачки: {error_msg}")

        # Проверяем, не связана ли ошибка с истечением ключа
        if check_api_key_error(error_msg):
            logger.critical("⛔️ ОБНАРУЖЕНА КРИТИЧЕСКАЯ ОШИБКА: Истек срок действия API-ключа!")
            raise APIKeyExpiredError(f"API-ключ истек: {error_msg[:200]}...")

        return [
            {
                "excel_row": data["excel_row"],
                "df_index": data["df_index"],
                "correctness": "Ошибка API",
                "error_type": f"Ошибка при обработке: {error_msg[:100]}"
            }
            for data in batch_data
        ]


# ===== ГЛАВНАЯ ФУНКЦИЯ =====
def analyze_report(input_excel_path, agent_name):
    """
    Основная функция анализа отчёта: загружает Excel, фильтрует диалоги,
    разбивает на пачки, отправляет в нейросеть и возвращает DataFrame
    с колонками «Корректность» и «Тип ошибки».

    Args:
        input_excel_path: путь к Excel-файлу с диалогами
        agent_name: имя агента из Config.AGENTS

    Returns:
        Кортеж (DataFrame с результатами, путь_к_лог_файлу) при успехе.
        (None, путь_к_файлу_ошибки) при ошибке.

    Raises:
        APIKeyExpiredError: если API-ключ истёк или недействителен.
        ValueError: если агент не найден или отсутствуют нужные колонки.
    """
    prompt_id = Config.AGENTS.get(agent_name)
    if not prompt_id:
        raise ValueError(f"❌ Агент '{agent_name}' не найден в конфигурации")

    # 2. СОЗДАЁМ ЛОГГЕР
    timestamp = datetime.now().strftime('%d.%m.%Y_%H-%M')
    agent_filename = agent_name.replace(" ", "_").replace("-", "_").lower()

    # Получаем имя исходного файла без расширения
    original_name = Path(input_excel_path).stem

    # Формируем единый префикс для всех файлов этого анализа
    file_prefix = f"{agent_filename}_{timestamp}_{original_name}"

    # Имена файлов
    log_filename = f"logs/{file_prefix}.log"
    error_filename = f"logs/ERROR_{file_prefix}.txt"

    Path("logs").mkdir(exist_ok=True)

    logger = logging.getLogger(f"analyzer_{timestamp}")
    logger.setLevel(logging.INFO)
    logger.propagate = False
    logger.handlers = []

    file_handler = logging.FileHandler(log_filename, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # 3. ТЕПЕРЬ можно логировать
    logger.info("=" * 60)
    logger.info(f"НАЧАЛО АНАЛИЗА: {timestamp}")
    logger.info(f"Агент: {agent_name}")
    logger.info(f"Prompt ID: {prompt_id[:10]}...")
    logger.info(f"Исходный файл: {input_excel_path}")
    logger.info(f"Лог-файл: {log_filename}")
    logger.info("=" * 60)

    try:
        # 3. ЗАГРУЖАЕМ EXCEL-ФАЙЛ
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

        # Логируем колонки только если их немного
        if len(df.columns) <= 20:
            logger.info(f"Колонки: {', '.join(df.columns.tolist())}")

        # 4. ПРОВЕРКА НЕОБХОДИМЫХ КОЛОНОК
        required_cols = ["Текст разговора с роботом", "results-confirmation"]
        missing_cols = [col for col in required_cols if col not in df.columns]

        if missing_cols:
            error_msg = f"❌ Отсутствуют колонки: {missing_cols}"
            logger.error(error_msg)
            raise ValueError(error_msg)

        logger.info("✅ Все необходимые колонки присутствуют")

        # 5. СОЗДАЁМ КЛИЕНТ ДЛЯ YANDEX CLOUD
        logger.info("Создаю клиент для Yandex Cloud...")
        client = openai.OpenAI(
            api_key=Config.YANDEX_API_KEY,
            base_url=Config.YANDEX_BASE_URL,
            project=Config.YANDEX_PROJECT_ID
        )
        logger.info("✅ Клиент создан")

        # 6. ПОДГОТОВКА ДАННЫХ
        logger.info("Подготовка данных...")

        # СОХРАНЯЕМ КОПИЮ ИСХОДНЫХ ДАННЫХ ДЛЯ ПОЛНОГО ОТЧЕТА (БУДЕТ НУЖНО ПОЗЖЕ)
        df_original = df.copy()

        # ФИЛЬТРАЦИЯ ДАННЫХ ДЛЯ AI-АНАЛИЗА
        logger.info("Фильтрация данных для AI-анализа...")
        df_filtered = filter_report(df, logger)

        # Создаём колонки если их нет
        if "Корректность" not in df_filtered.columns:
            df_filtered["Корректность"] = ""
        if "Тип ошибки" not in df_filtered.columns:
            df_filtered["Тип ошибки"] = ""

        # Заполняем пустые значения
        df_filtered["Корректность"] = df_filtered["Корректность"].fillna("").astype(str)
        df_filtered["Тип ошибки"] = df_filtered["Тип ошибки"].fillna("").astype(str)

        # Подготавливаем список звонков для обработки (ИСПОЛЬЗУЕМ ОТФИЛЬТРОВАННЫЕ ДАННЫЕ)
        calls_to_process = []
        for idx, row in df_filtered.iterrows():
            dialogue = str(row["Текст разговора с роботом"]) if pd.notna(row["Текст разговора с роботом"]) else ""

            # Используем существующий ID или создаём новый
            call_id = row.get("Идентификатор звонка", f"Строка_{idx + 2}")

            calls_to_process.append({
                "excel_row": idx + 2,  # +2 потому что Excel нумерует с 1 + заголовок
                "df_index": idx,  # Индекс в DataFrame (начинается с 0)
                "call_id": call_id,
                "dialogue": dialogue,
                "result": str(row["results-confirmation"]) if pd.notna(row["results-confirmation"]) else ""
            })

        logger.info(f"📊 Подготовлено {len(calls_to_process)} диалогов для анализа")

        # 7. РАЗБИВАЕМ НА ПАЧКИ И АНАЛИЗИРУЕМ
        batch_size = 25
        batches = []
        for i in range(0, len(calls_to_process), batch_size):
            batch = calls_to_process[i:i + batch_size]
            if batch:
                batches.append(batch)

        logger.info(f"📦 Создано {len(batches)} пачек по {batch_size} диалогов")

        total_processed = 0
        correct_count = 0
        start_time = time.time()

        # 8. ОБРАБОТКА КАЖДОЙ ПАЧКИ
        for batch_num, batch in enumerate(batches, 1):
            logger.info(f"Обработка пачки {batch_num}/{len(batches)} ({len(batch)} диалогов)...")

            batch_results = analyze_batch(batch, logger, client, prompt_id)

            # Записываем результаты в DataFrame (ОТФИЛЬТРОВАННЫЙ)
            for result in batch_results:
                if 0 <= result["df_index"] < len(df_filtered):
                    df_filtered.at[result["df_index"], "Корректность"] = result["correctness"]
                    df_filtered.at[result["df_index"], "Тип ошибки"] = result["error_type"]

                    if result["correctness"] == "Корректно":
                        correct_count += 1
                else:
                    logger.error(f"Неверный индекс {result['df_index']} для строки {result['excel_row']}")

            total_processed += len(batch)

            # Пауза между пачками
            if batch_num < len(batches):
                time.sleep(1)

        # ПЕРЕУПОРЯДОЧИВАЕМ КОЛОНКИ
        if "Корректность" in df_filtered.columns and "Тип ошибки" in df_filtered.columns:
            # Создаем новый порядок колонок
            new_order = []
            for col in df_filtered.columns:
                if col == "Текст разговора с роботом":
                    new_order.extend([col, "Корректность", "Тип ошибки"])
                elif col not in ["Корректность", "Тип ошибки"]:
                    new_order.append(col)

            # Применяем новый порядок
            df_filtered = df_filtered[new_order]
            logger.info(
                f"✅ Колонки переупорядочены: 'Корректность' и 'Тип ошибки' перемещены после 'Текст разговора с роботом'")

        # ВМЕСТО СОХРАНЕНИЯ В ФАЙЛ - ВОЗВРАЩАЕМ DataFrame
        logger.info(f"✅ Анализ завершён. Результаты готовы ({len(df_filtered)} строк)")

        # 10. СТАТИСТИКА
        total_time = time.time() - start_time
        original_count = len(df_original)
        filtered_count = len(df_filtered)
        filtered_pct = f"{filtered_count / original_count * 100:.1f}%" if original_count else "—"
        correct_pct = f"{correct_count / total_processed * 100:.1f}%" if total_processed else "—"

        logger.info("=" * 60)
        logger.info("📊 ИТОГИ АНАЛИЗА:")
        logger.info(f"   Агент: {agent_name}")
        logger.info(f"   Исходных диалогов: {original_count}")
        logger.info(f"   После фильтрации: {filtered_count} ({filtered_pct})")
        logger.info(f"   Проанализировано: {total_processed}")
        logger.info(f"   Корректных: {correct_count} ({correct_pct})")
        logger.info(f"   Время обработки: {total_time:.1f} секунд")
        logger.info(f"   Лог-файл: {log_filename}")
        logger.info("=" * 60)

        logger.info("✅ Анализ завершён успешно! Возвращаю DataFrame.")

        # ВОЗВРАЩАЕМ DataFrame И ПУТЬ К ЛОГУ
        return df_filtered, log_filename

    except APIKeyExpiredError as e:
        # Логируем критическую ошибку истечения ключа
        logger.critical("=" * 60)
        logger.critical("⛔️ АНАЛИЗ ПРЕРВАН ИЗ-ЗА ИСТЕЧЕНИЯ API-КЛЮЧА!")
        logger.critical(f"Ошибка: {str(e)}")
        logger.critical("=" * 60)

        # Бросаем исключение дальше, чтобы бот мог его поймать и остановить
        raise APIKeyExpiredError(f"API-ключ истек: {str(e)[:200]}...")

    except Exception as e:
        # Если произошла другая ошибка - логируем
        logger.error("❌ КРИТИЧЕСКАЯ ОШИБКА:")
        logger.error(f"   Тип ошибки: {type(e).__name__}")
        logger.error(f"   Сообщение: {str(e)}")
        logger.error("=" * 60)

        # Создаём файл с подробной информацией об ошибке
        with open(error_filename, 'w', encoding='utf-8') as f:
            f.write(f"Ошибка при анализе {input_excel_path}:\n")
            f.write(f"Время: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
            f.write(f"Агент: {agent_name}\n")
            f.write(f"Тип ошибки: {type(e).__name__}\n")
            f.write(f"Сообщение: {str(e)}\n\n")
            f.write("Трассировка:\n")
            f.write(traceback.format_exc())

        # Возвращаем None для файла результата и лог с ошибкой
        return None, error_filename

    finally:
        logger.removeHandler(file_handler)
        file_handler.close()
