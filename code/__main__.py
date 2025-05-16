from processor.base import ExcelProcessor
from config import LOG_FILE

if __name__ == "__main__":
    processor = ExcelProcessor()
    processor.process_all()
    print(f"\nProcessing complete! Check log file for details: {LOG_FILE}")
