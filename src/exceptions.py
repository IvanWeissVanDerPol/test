class ExcelProcessorException(Exception):
    """Base exception for Excel processing errors."""
    pass

class InvalidFilePathError(ExcelProcessorException):
    """Raised when an invalid file path is provided."""
    def __init__(self, path: str):
        super().__init__(f"Invalid file path: {path}")

class MissingColumnError(ExcelProcessorException):
    """Raised when a required column is missing."""
    def __init__(self, column_name: str):
        super().__init__(f"Required column not found: {column_name}")

class InvalidTimeFormatError(ExcelProcessorException):
    """Raised when time format is invalid."""
    def __init__(self, time_str: str):
        super().__init__(f"Invalid time format: {time_str}")
