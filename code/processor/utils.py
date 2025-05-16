from datetime import datetime, time, timedelta
import pandas as pd
from typing import Optional, Dict, Any, List, Union, Tuple
import inspect
import re
from logger_config import setup_logger

logger = setup_logger(__name__)

def log_variables(local_vars: Dict[str, Any], exclude: Optional[List[str]] = None) -> None:
    """Log variable names and their values"""
    if exclude is None:
        exclude = []
    exclude.extend(['self', 'args', 'kwargs', 'exclude'])
    
    # Skip debug logging in production
    if logger.level > 10:  # DEBUG level is 10
        return
        
    frame = inspect.currentframe().f_back
    try:
        for var_name, var_value in frame.f_locals.items():
            if var_name not in exclude:
                logger.debug(f"Variable: {var_name} = {var_value!r}")
    finally:
        del frame

def is_formula(value: Any) -> bool:
    """Check if the value is an Excel formula"""
    return isinstance(value, str) and value.startswith('=')

def parse_decimal_time(decimal_hours: Union[str, float, int]) -> Optional[time]:
    """
    Convert decimal hours to time object (e.g., 1.5 -> 01:30:00)
    
    Args:
        decimal_hours: Time in decimal hours (e.g., 1.5 for 1 hour 30 minutes)
        
    Returns:
        time object or None if conversion fails
    """
    try:
        if pd.isna(decimal_hours) or decimal_hours == '':
            return None
            
        # Convert to float if it's a string
        if isinstance(decimal_hours, str):
            decimal_hours = float(decimal_hours.replace(',', '.'))
            
        # Handle negative values (if needed)
        is_negative = decimal_hours < 0
        decimal_hours = abs(decimal_hours)
        
        # Handle values >= 24 hours by taking modulo 24
        decimal_hours = decimal_hours % 24
        
        hours = int(decimal_hours)
        minutes = int((decimal_hours - hours) * 60)
        seconds = int((((decimal_hours - hours) * 60) - minutes) * 60)
        
        # Handle overflow
        if seconds >= 60:
            minutes += seconds // 60
            seconds = seconds % 60
        if minutes >= 60:
            hours += minutes // 60
            minutes = minutes % 60
            
        return time(hour=hours, minute=minutes, second=seconds)
        
    except (ValueError, TypeError) as e:
        logger.warning(f"Could not convert {decimal_hours} to time: {e}")
        return None

def format_datetime(value):
    """Format datetime value to standard string format"""
    log_variables(locals())
    
    if pd.isna(value):
        logger.debug("Input value is NA/NaN, returning None")
        return None
        
    try:
        if isinstance(value, str):
            logger.debug(f"Formatting datetime string: {value}")
            formats = [
                '%d/%m/%Y %H:%M:%S',  # 31/12/2022 23:59:59
                '%Y-%m-%d %H:%M:%S',  # 2022-12-31 23:59:59
                '%d-%m-%Y %H:%M:%S',  # 31-12-2022 23:59:59
                '%d/%m/%Y %H:%M',     # 31/12/2022 23:59
                '%Y-%m-%d %H:%M',     # 2022-12-31 23:59
                '%d-%m-%Y %H:%M'      # 31-12-2022 23:59
            ]
            
            for fmt in formats:
                try:
                    dt = datetime.strptime(value, fmt)
                    result = dt.strftime('%d/%m/%Y %H:%M:%S')
                    logger.debug(f"Successfully parsed with format '{fmt}': {value} -> {result}")
                    return result
                except ValueError:
                    continue
            
            logger.warning(f"Could not parse datetime string with any known format: {value}")
            return value
        else:
            logger.debug(f"Converting non-string value to datetime: {value}")
            try:
                dt = pd.to_datetime(value)
                result = dt.strftime('%d/%m/%Y %H:%M:%S')
                logger.debug(f"Converted to datetime: {value} -> {result}")
                return result
            except Exception as e:
                logger.error(f"Error converting value to datetime: {value}, error: {e}")
                return value
    except Exception as e:
        logger.error(f"Unexpected error in format_datetime: {e}", exc_info=True)
        return value

def parse_time_string(time_val: Union[str, float, int]) -> Optional[time]:
    """
    Parse a time value into a time object.
    Handles Excel formulas, decimal hours, and various time string formats.
    
    Args:
        time_val: Time value to parse (string, float, or int)
        
    Returns:
        time object or None if parsing fails
    """
    # Skip logging for performance in production
    if logger.level <= 10:  # DEBUG level
        log_variables(locals())
    
    # Handle None/NA/empty values
    if not time_val or pd.isna(time_val) or str(time_val).strip() == '':
        return None
    
    # Skip Excel formulas
    if is_formula(time_val):
        logger.debug(f"Skipping formula cell: {time_val}")
        return None
    
    try:
        # Try parsing as decimal time first (e.g., 1.5 for 1:30)
        if isinstance(time_val, (int, float)) or \
           (isinstance(time_val, str) and re.match(r'^\s*\d+([.,]\d+)?\s*$', time_val)):
            if isinstance(time_val, str):
                # Normalize decimal separator
                time_val = time_val.replace(',', '.')
            return parse_decimal_time(time_val)
        
        # Handle string time values
        if isinstance(time_val, str):
            cleaned_time = time_val.strip()
            
            # Try standard time formats
            time_formats = [
                '%H:%M',        # 14:30
                '%H:%M:%S',     # 14:30:45
                '%I:%M %p',     # 2:30 PM
                '%I:%M%p',      # 2:30PM
                '%I.%M %p',     # 2.30 PM
                '%I.%M%p',      # 2.30PM
                '%H%M',         # 1430
                '%H%M%S',       # 143045
                '%H.%M',        # 14.30
                '%H.%M.%S'      # 14.30.45
            ]
            
            for fmt in time_formats:
                try:
                    time_obj = datetime.strptime(cleaned_time, fmt).time()
                    logger.debug(f"Parsed time: {time_val} -> {time_obj} (format: {fmt})")
                    return time_obj
                except ValueError:
                    continue
        
        logger.warning(f"Could not parse time value: {time_val}")
        return None
        
    except Exception as e:
        logger.error(f"Error parsing time value '{time_val}': {str(e)}")
        return None

def process_time_cells(worksheet, time_columns: List[int], rows_to_process: List[int]) -> Dict[str, int]:
    """
    Process time cells in the specified columns and rows
    
    Args:
        worksheet: OpenPyXL worksheet object
        time_columns: List of 1-based column indices to process
        rows_to_process: List of 1-based row indices to process
        
    Returns:
        Dictionary with processing statistics
    """
    results = {
        'processed': 0,
        'skipped_formulas': 0,
        'errors': 0,
        'unchanged': 0
    }
    
    start_time = datetime.now()
    logger.info(f"Starting time processing for {len(time_columns)} columns and {len(rows_to_process)} rows")
    
    try:
        for col_idx in time_columns:
            col_letter = chr(64 + col_idx) if col_idx <= 26 else chr(64 + (col_idx-1)//26) + chr(65 + (col_idx-1)%26)
            
            for row_idx in rows_to_process:
                cell = worksheet.cell(row=row_idx, column=col_idx)
                
                # Skip empty cells
                if cell.value is None or cell.value == '':
                    results['unchanged'] += 1
                    continue
                    
                # Skip formula cells
                if is_formula(cell.value):
                    results['skipped_formulas'] += 1
                    continue
                
                try:
                    # Parse the time value
                    time_val = cell.value
                    time_obj = parse_time_string(time_val)
                    
                    if time_obj is not None:
                        # Only update if the value has changed
                        if not (isinstance(cell.value, time) and cell.value == time_obj):
                            cell.value = time_obj
                            results['processed'] += 1
                        else:
                            results['unchanged'] += 1
                    else:
                        results['unchanged'] += 1
                        
                except Exception as e:
                    results['errors'] += 1
                    logger.error(f"Error processing cell {col_letter}{row_idx}: {str(e)}")
                    
    except Exception as e:
        logger.error(f"Unexpected error during time processing: {str(e)}", exc_info=True)
        results['errors'] += 1
    
    duration = (datetime.now() - start_time).total_seconds()
    logger.info(
        f"Time processing completed in {duration:.2f} seconds. "
        f"Processed: {results['processed']}, "
        f"Skipped formulas: {results['skipped_formulas']}, "
        f"Unchanged: {results['unchanged']}, "
        f"Errors: {results['errors']}"
    )
    
    return results
