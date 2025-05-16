import os
from pathlib import Path
from datetime import datetime

# Project root directory
PROJECT_ROOT = Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Directories
DIRECTORIES = {
    'data': PROJECT_ROOT / 'data',
    'logs': PROJECT_ROOT / 'logs',
    'output': PROJECT_ROOT / 'output',
    'input': PROJECT_ROOT / 'data/input'
}

# file {file name and path}
FILE_PATHS = {
    'input': {
        'daily_file': DIRECTORIES['input'] / 'A-INFORME QUÍMICO DIARIO 2025 Macro prueba.xlsx',
        'data_ambar_original': DIRECTORIES['data'] / 'Data_Ambar Macro Prueba.xlsx'
    },
    'output': {
        'daily_file_cleaned': DIRECTORIES['output'] / 'A-INFORME QUÍMICO DIARIO 2025 Macro prueba_updated.xlsx',
        'daily_file_filtered': DIRECTORIES['output'] / 'A-INFORME QUÍMICO DIARIO 2025 Macro prueba_filtered.xlsx',
        'data_ambar_updated': DIRECTORIES['output'] / 'Data_Ambar Macro Prueba_updated.xlsx'
    },
    'log_file': DIRECTORIES['logs'] / 'pipeline.log'
}

# Create directories if they don't exist
for dir_path in DIRECTORIES.values():
    dir_path.mkdir(parents=True, exist_ok=True)


VALUES_TO_SET = [
    "Semillas L593",
    "Semillas L594",
    "Semillas (0 - 0,5) mm L 593",
    "Semillas (0 - 0,5) mm L 594",
    "Burbujas ( 0,5-1) mm L 593",
    "Burbujas ( 0,5-1) mm L 594",
    "Burbujas ( >1)mm L 593",
    "Burbujas ( >1)mm L 594"
]

# Columns configuration
COLUMN_CONFIG = {
    'columns_to_keep': [
        "Hora de Análisis",
        "Saturación (%) (Pureza)",
        "Longitud de onda (nm)",
        "L*",
        "a*",
        "b*",
        "Densidad",
        "% T 550 (2mm)",
        "Semillas L593",
        "Semillas L594",
        "Semillas (0 - 0,5) mm L 593",
        "Semillas (0 - 0,5) mm L 594",
        "Burbujas ( 0,5-1) mm L 593",
        "Burbujas ( 0,5-1) mm L 594",
        "Burbujas ( >1)mm L 593",
        "Burbujas ( >1)mm L 594",
        "Burbujas por Kg - 593",
        "Burbujas por Kg - 594",
        "SiO2",
        "Na2O",
        "CaO",
        "MgO",
        "Al2O3",
        "K2O",
        "SO3",
        "Fe2O3",
        "TiO2",
        "SiO2D (100-S(ox))",
        "Cr2O3",
        "%FeO as Fe2O3",
        "Redox",
        "Viscosidad (°C)",
        "Cooling Time (s)"
    ],
    'date_columns': ['FECHA'],
    'column_mapping': {
        'Hora de Análisis': 'Hora de Análisis',
        'Saturación (%) (Pureza)': 'Pureza',
        'Longitud de onda (nm)': 'DWL',
        'L*': 'L',
        'a*': 'a',
        'b*': 'b',
        'Densidad': 'Densidad',
        '% T 550 (2mm)': '%T 550nm (2mm)',
        'Semillas L593': 'Semillas L593',
        'Semillas L594': 'Semillas L594',
        'Semillas (0 - 0,5) mm L 593': 'Semillas (0 - 0,5) mm L 593',
        'Semillas (0 - 0,5) mm L 594': 'Semillas (0 - 0,5) mm L 594',
        'Burbujas (0,5-1) mm L 593': 'Burbujas (0,5-1) mm L 593',
        'Burbujas (0,5-1) mm L 594': 'Burbujas (0,5-1) mm L 594',
        'Burbujas (>1)mm L 593': 'Burbujas (>1)mm L 593',
        'Burbujas (>1)mm L 594': 'Burbujas (>1)mm L 594',
        'Burbujas por Kg - 593': 'Burbujas L993/kg',
        'Burbujas por Kg - 594': 'Burbujas L994/kg',
        'SiO2': 'SiO2',
        'Na2O': 'Na2O',
        'CaO': 'CaO',
        'MgO': 'MgO',
        'Al2O3': 'Al2O3',
        'K2O': 'K2O',
        'SO3': 'SO3',
        'Fe2O3': 'Fe2O3',
        'TiO2': 'TiO2',
        'SiO2D (100-S(ox))': 'SiO2D (100-S(ox))',
        'Cr2O3': 'Cr2O3',
        '%FeO as Fe2O3': 'FeO',
        'Redox': 'Redox',
        'Viscosidad (°C)': 'Viscosidad (°C)',
        'Cooling Time (s)': 'Cooling Time (s)'
    }
}
