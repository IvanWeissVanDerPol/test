import os
from datetime import datetime

BASE_DIR = r"c:\Users\weiss\Desktop\test"
DATA_DIR = os.path.join(BASE_DIR, "data")
ORIGINAL_FILE = os.path.join(DATA_DIR, "A-INFORME QUÍMICO DIARIO 2025 Macro prueba.xlsx")
UPDATED_FILE = os.path.splitext(ORIGINAL_FILE)[0] + '_updated.xlsx'
FILTERED_FILE = os.path.splitext(ORIGINAL_FILE)[0] + '_updated_filtered.xlsx'
LOG_FILE = f'data_processor_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
