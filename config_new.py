"""
Configuration Module
Berisi semua konstanta dan environment variables
"""
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# ===== API KEYS =====
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "openai/gpt-4o-mini")
SERPER_API_KEY = os.getenv("SERPER_API_KEY")

# ===== DIRECTORIES =====
BASE_DIR = Path(__file__).parent
FILES_DIR = BASE_DIR / "static" / "files"
TEMPLATE_FILE = BASE_DIR / "template_quotation.docx"
TEMP_DIR = BASE_DIR / "temp"
COUNTER_FILE = BASE_DIR / "counter.json"
DB_FILE = BASE_DIR / "chat_history.db"

# Ensure directories exist
FILES_DIR.mkdir(parents=True, exist_ok=True)
TEMP_DIR.mkdir(parents=True, exist_ok=True)

# ===== FLASK CONFIG =====
FLASK_SECRET_KEY = "karya-limbah-2025"
FLASK_PORT = int(os.getenv("PORT", 5000))
FLASK_DEBUG = os.getenv("FLASK_ENV") != "production"

# ===== BULAN INDONESIA =====
BULAN_ID = {
    '01': 'Januari', '02': 'Februari', '03': 'Maret', '04': 'April',
    '05': 'Mei', '06': 'Juni', '07': 'Juli', '08': 'Agustus',
    '09': 'September', '10': 'Oktober', '11': 'November', '12': 'Desember'
}