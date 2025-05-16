import pandas as pd
import numpy as np
import os
import shutil
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
import logging
from src.utils.date_utils import format_datetime

# Set up logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Add file handler for logging to file
date_str = datetime.now().strftime('%Y%m%d_%H%M%S')
log_file = f'transfer_data_{date_str}.log'
file_handler = logging.FileHandler(log_file)
file_handler.setLevel(logging.DEBUG)
formatter = logging.Formatter('%%(asctime)s - %%(levelname)s - %%(message)s')
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

def transfer_data():
    try:
        logger.info("Starting data transfer process")
        
        # Define the mapping between columns
        column_mapping = {
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
        logger.debug(f"Column mapping: {column_mapping}")

        from src.config.config import FILE_PATHS
        
        # Load the source file using config path
        source_path = str(FILE_PATHS['input']['data_ambar_original'])
        logger.info(f"Loading source file: {source_path}")
        
        # Create updated path using config
        updated_path = str(FILE_PATHS['output']['data_ambar_updated'])
        logger.info(f"Creating copy at: {updated_path}")
        shutil.copy2(source_path, updated_path)

        # Load filtered file using config path
        filtered_path = str(FILE_PATHS['output']['daily_file_filtered'])
        logger.info(f"Loading filtered file: {filtered_path}")
        filtered_df = pd.read_excel(filtered_path)
        logger.info(f"Filtered file loaded with {len(filtered_df)} rows")
        
        # Read the copied file and find the correct sheet name
        with pd.ExcelFile(updated_path) as xls:
            sheet_names = xls.sheet_names
            logger.debug(f"Available sheet names: {sheet_names}")
            
            # Try different variations of 'datos'
            found = False
            for sheet in sheet_names:
                if sheet.lower() in ['datos', 'data', 'datos sheet', 'hoja datos']:
                    logger.info(f"Found matching sheet: {sheet}")
                    updated_df = pd.read_excel(updated_path, sheet_name=sheet)
                    found = True
                    break
            
            if not found:
                logger.warning(f"No matching sheet found, using first sheet: {sheet_names[0]}")
                updated_df = pd.read_excel(updated_path, sheet_name=sheet_names[0])

        # Drop empty rows
        original_rows = len(updated_df)
        updated_df = updated_df.dropna(how='all')
        dropped_rows = original_rows - len(updated_df)
        logger.info(f"Dropped {dropped_rows} completely empty rows")

        # Add data from filtered file to the copied file
        logger.info("Adding data from filtered file...")
        for _, row in filtered_df.iterrows():
            new_row = {}
            for target_col_name, source_col_name in column_mapping.items():
                if target_col_name in row:
                    value = row[target_col_name]
                    if target_col_name == 'Hora de Análisis':
                        value = format_datetime(value)
                    logger.debug(f"Mapping {target_col_name} -> {source_col_name}: {value}")
                    new_row[source_col_name] = value

            # Create a DataFrame from the new row and append it
            new_row_df = pd.DataFrame([new_row])
            updated_df = pd.concat([updated_df, new_row_df], ignore_index=True)
        logger.info(f"Added {len(filtered_df)} rows from filtered file")

        # Save the updated file with pandas while trying to preserve formatting
        try:
            # Use pandas for saving since openpyxl has issues with pivot tables
            logger.info("Using pandas for saving to avoid pivot table issues...")
            
            # First, get the sheet name from the original file
            with pd.ExcelFile(updated_path) as xls:
                sheet_names = xls.sheet_names
                logger.debug(f"Available sheet names: {sheet_names}")
                
                # Try different variations of 'datos'
                sheet_name = None
                for sheet in sheet_names:
                    if sheet.lower() in ['datos', 'data', 'datos sheet', 'hoja datos']:
                        sheet_name = sheet
                        logger.info(f"Using sheet name: {sheet_name}")
                        break
                else:
                    # If no matching sheet found, use the first sheet
                    sheet_name = sheet_names[0]
                    logger.warning(f"Using first sheet: {sheet_name}")
            
            # Save with pandas
            with pd.ExcelWriter(updated_path, mode='w', engine='openpyxl') as writer:
                updated_df.to_excel(writer, sheet_name=sheet_name, index=False)
            logger.info(f"Successfully saved with pandas to sheet: {sheet_name}")
        except Exception as e:
            logger.error(f"Error saving with pandas: {str(e)}")
            raise RuntimeError(f"Failed to save updated file: {str(e)}") from e

        logger.info(f"Data transfer complete!")
        logger.info(f"Updated file saved as: {updated_path}")
        logger.info(f"Final number of rows: {len(updated_df)}")
        
        print(f"\nData transfer complete!")
        print(f"Updated file saved as: {updated_path}")
        print(f"Check log file for detailed information: {log_file}")
        
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}", exc_info=True)
        print(f"Error occurred: {str(e)}")
        
        # Try to save with pandas as a last resort
        try:
            logger.info("Attempting final save with pandas...")
            with pd.ExcelWriter(updated_path, mode='w', engine='openpyxl') as writer:
                updated_df.to_excel(writer, sheet_name='datos', index=False)
            logger.info("Successfully saved with pandas as final fallback")
            print("Successfully saved with pandas as final fallback")
        except Exception as e:
            logger.error(f"Failed final save with pandas: {str(e)}", exc_info=True)
            print(f"Failed final save with pandas: {str(e)}")
            raise

if __name__ == "__main__":
    transfer_data()
