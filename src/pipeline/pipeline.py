from pathlib import Path
from typing import Optional
from src.excel.excel_processor import ExcelProcessor
from src.config.config import FILE_PATHS
import logging

logger = logging.getLogger(__name__)

class Pipeline:
    """
    Main pipeline class for processing Excel files.
    
    Attributes:
        processor: Current ExcelProcessor instance
        current_file: Current file being processed
        steps: List of completed pipeline steps
    """
    
    def __init__(self):
        """Initialize pipeline with empty state."""
        self.processor: Optional[ExcelProcessor] = None
        self.current_file: Optional[Path] = None
        self.steps: list[str] = []

    def clean_daily(self) -> Path:
        """
        Clean daily Excel file and update dates.
        
        Returns:
            Path to the cleaned file
            
        Raises:
            ValueError: If file cleaning fails
        """
        try:
            logger.info("Starting daily file cleaning")
            self.processor = ExcelProcessor(str(FILE_PATHS['input']['daily_file']))
            self.processor.update_dates()
            
            updated_path = Path(self.processor.file_path)
            self.current_file = updated_path
            self.steps.append('clean_daily')
            logger.info(f"Daily file cleaning completed. Saved as: {updated_path}")
            return updated_path
        except Exception as e:
            logger.error(f"Error in clean_daily: {str(e)}")
            raise ValueError(f"Failed to clean daily file: {str(e)}") from e

    def filter_columns(self, file_path: Path) -> Path:
        """Filter columns to keep only the specified ones."""
        logger.info("Starting column filtering")
        self.processor = ExcelProcessor(str(file_path))
        self.processor.filter_columns("Transposed")
        
        # Get the output path from the ExcelProcessor
        filtered_path = self.processor.file_path
        logger.info(f"Column filtering completed. Saved as: {filtered_path}")
        return filtered_path

    def transfer_data(self, file_path: Path) -> Path:
        """Transfer data to target file."""
        logger.info("Starting data transfer")
        self.processor = ExcelProcessor(str(file_path))
        self.processor.transfer_data()
        
        # Get the output path from the ExcelProcessor
        target_path = self.processor.file_path
        logger.info(f"Data transfer completed. Saved to: {target_path}")
        return target_path

    def run(self) -> Path:
        """
        Run the complete pipeline.
        
        Returns:
            Path to the final processed file
            
        Raises:
            RuntimeError: If pipeline fails at any step
        """
        try:
            # Reset pipeline state
            self.processor = None
            self.current_file = None
            self.steps = []
            
            # Step 1: Clean daily file
            # logger.info("Starting pipeline execution")
            # updated_file = self.clean_daily()
            
            #load the saved file from the output folder for testing pourpuses C:\Users\weiss\Desktop\test\src\output\processed_A-INFORME QUÍMICO DIARIO 2025 Macro prueba.xlsx
            updated_file = Path(r"C:\Users\weiss\Desktop\test\src\output\processed_A-INFORME QUÍMICO DIARIO 2025 Macro prueba.xlsx")
            
            # Step 2: Filter columns
            filtered_file = self.filter_columns(updated_file)
            
            # Step 3: Transfer data
            target_file = self.transfer_data(filtered_file)
            
            logger.info(f"Pipeline completed successfully. Final file: {target_file}")
            return target_file
            
        except Exception as e:
            logger.error(f"Pipeline failed: {str(e)}")
            logger.error(f"Completed steps: {self.steps}")
            raise RuntimeError(f"Pipeline failed at step: {self.steps[-1] if self.steps else 'initialization'}") from e
