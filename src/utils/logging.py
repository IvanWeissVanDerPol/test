import logging
from src.config.config import FILE_PATHS

def setup_logging():
    """Setup logging configuration."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(FILE_PATHS['log_file']),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)
