import pandas as pd
import os
import shutil
from typing import Dict, Any, List, Optional
import inspect

from config import BASE_DIR, FILTERED_FILE
from logger_config import setup_logger
from processor.utils import format_datetime

logger = setup_logger(__name__)

def log_variables(local_vars: Dict[str, Any], exclude: Optional[List[str]] = None) -> None:
    """Log variable names and their values"""
    if exclude is None:
        exclude = []
    exclude.extend(['self', 'args', 'kwargs', 'exclude'])
    
    frame = inspect.currentframe().f_back
    try:
        for var_name, var_value in frame.f_locals.items():
            if var_name not in exclude:
                logger.debug(f"Variable: {var_name} = {var_value!r}")
    finally:
        del frame

def transfer_data() -> None:
    logger.info("Starting data transfer from filtered to source file...")
    log_variables(locals())
    
    try:
        # Define column mapping with detailed logging
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
        
        # Set up source and destination paths
        source_path = os.path.join(BASE_DIR, "data", "Data_Ambar Macro Prueba.xlsx")
        updated_path = os.path.splitext(source_path)[0] + '_updated.xlsx'
        logger.info(f"Source file: {source_path}")
        logger.info(f"Destination file: {updated_path}")
        
        # Create a backup of the source file
        logger.info("Creating backup of source file...")
        shutil.copy2(source_path, updated_path)
        logger.info(f"Backup created at: {updated_path}")

        # Read filtered data
        logger.info(f"Reading filtered data from: {FILTERED_FILE}")
        filtered_df = pd.read_excel(FILTERED_FILE)
        logger.info(f"Read {len(filtered_df)} rows from filtered data")
        
        # Find the correct sheet in the destination file
        logger.info("Searching for the correct sheet in the destination file...")
        with pd.ExcelFile(updated_path) as xls:
            sheet_found = False
            for sheet in xls.sheet_names:
                if sheet.lower() in ['datos', 'data', 'datos sheet', 'hoja datos']:
                    logger.info(f"Found target sheet: {sheet}")
                    updated_df = pd.read_excel(updated_path, sheet_name=sheet)
                    sheet_found = True
                    break
            
            if not sheet_found:
                logger.warning("No matching sheet found, using first sheet")
                updated_df = pd.read_excel(updated_path)
        
        logger.info(f"Original data shape before processing: {updated_df.shape}")
        
        # Clean up the dataframe
        initial_rows = len(updated_df)
        updated_df = updated_df.dropna(how='all')
        logger.info(f"Dropped {initial_rows - len(updated_df)} empty rows from original data")

        # Process each row in the filtered data
        logger.info("Processing filtered data rows...")
        rows_added = 0
        for idx, row in filtered_df.iterrows():
            new_row = {}
            logger.debug(f"Processing row {idx + 1}/{len(filtered_df)}")
            
            for target_col, source_col in column_mapping.items():
                if target_col in row:
                    val = row[target_col]
                    if target_col == 'Hora de Análisis':
                        logger.debug(f"Formatting datetime for column: {target_col}")
                        formatted_val = format_datetime(val)
                        logger.debug(f"Formatted datetime: {val} -> {formatted_val}")
                        val = formatted_val
                    new_row[source_col] = val
            
            logger.debug(f"Adding new row with data: {new_row}")
            updated_df = pd.concat([updated_df, pd.DataFrame([new_row])], ignore_index=True)
            rows_added += 1
        
        logger.info(f"Added {rows_added} new rows to the data")
        logger.info(f"Final data shape: {updated_df.shape}")

        # Save the updated data back to Excel
        logger.info(f"Saving updated data to: {updated_path}")
        with pd.ExcelWriter(updated_path, engine='openpyxl', mode='w') as writer:
            updated_df.to_excel(writer, sheet_name='datos', index=False)
        
        logger.info(f"Data transfer completed successfully. File saved at: {updated_path}")
        
    except Exception as e:
        logger.error(f"Error in transfer_data: {str(e)}", exc_info=True)
        raise
