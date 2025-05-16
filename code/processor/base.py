from processor.cleaning import clean_daily_excel
from processor.filtering import filter_columns
from processor.transferring import transfer_data
from logger_config import setup_logger
import inspect

logger = setup_logger(__name__)

def log_variables(local_vars):
    """Log variable names and their values"""
    frame = inspect.currentframe().f_back
    try:
        for var_name, var_value in frame.f_locals.items():
            if var_name not in ['self', 'args', 'kwargs']:
                logger.debug(f"Variable: {var_name} = {var_value!r}")
    finally:
        del frame

class ExcelProcessor:
    def process_all(self):
        try:
            logger.info("Starting complete processing sequence...")
            log_variables(locals())
            
            logger.info("Executing clean_daily_excel()")
            clean_daily_excel()
            
            logger.info("Executing filter_columns()")
            filter_columns()
            
            logger.info("Executing transfer_data()")
            transfer_data()
            
            logger.info("Processing completed successfully!")
        except Exception as e:
            logger.error(f"Error in process_all: {str(e)}", exc_info=True)
            raise
