import os
import sys
from pathlib import Path
from dotenv import load_dotenv

# Определяем базовую папку в зависимости от способа запуска
if getattr(sys, 'frozen', False):
    # Запуск из собранного exe
    if hasattr(sys, '_MEIPASS'):
        # Режим --onefile: временная папка с распакованными файлами
        BASE_DIR = Path(sys._MEIPASS)
    else:
        # Режим --onedir: папка с exe
        BASE_DIR = Path(sys.executable).parent
else:
    # Обычный запуск скрипта (из исходников)
    BASE_DIR = Path(__file__).parent

# Загружаем .env из этой папки
dotenv_path = BASE_DIR / '.env'
load_dotenv(dotenv_path=dotenv_path)


class Config:
    # Данные Yandex Cloud из переменных окружений
    YANDEX_API_KEY = os.getenv("YANDEX_API_KEY")
    YANDEX_BASE_URL = os.getenv("YANDEX_BASE_URL")
    YANDEX_PROJECT_ID = os.getenv("YANDEX_PROJECT_ID")

    # ID всех агентов в AI Studio
    AGENTS = {
        "Outbound assistant": os.getenv("YANDEX_AGENT_OUTBOUND_ASSISTANT"),
        "Predictive dialer": os.getenv("YANDEX_AGENT_PREDICTIVE_DIALER"),
        "Inbound assistant": os.getenv("YANDEX_AGENT_INBOUND_ASSISTANT"),
        "Deals bot": os.getenv("YANDEX_AGENT_DEALS_BOT"),
        "Reactivation bot": os.getenv("YANDEX_AGENT_REACTIVATION_BOT"),
        "Inbound support": os.getenv("YANDEX_AGENT_INBOUND_SUPPORT"),
        "Consultant bot": os.getenv("YANDEX_AGENT_CONSULTANT_BOT"),
        "Leads bot": os.getenv("YANDEX_AGENT_LEADS_BOT"),
        "Finance bot": os.getenv("YANDEX_AGENT_FINANCE_BOT"),
        "Wholesale bot": os.getenv("YANDEX_AGENT_WHOLESALE_BOT"),
        "Cold calls bot": os.getenv("YANDEX_AGENT_COLD_CALLS_BOT"),
        "Insurance bot": os.getenv("YANDEX_AGENT_INSURANCE_BOT"),
    }

    TEST_PHONES = []
    test_phones_str = os.getenv("TEST_PHONES", "")
    if test_phones_str:
        TEST_PHONES = [p.strip() for p in test_phones_str.split(",") if p.strip()]