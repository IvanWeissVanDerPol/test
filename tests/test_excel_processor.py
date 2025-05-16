import unittest
from unittest.mock import MagicMock, patch
from src.excel.excel_processor import ExcelProcessor
from src.exceptions import InvalidFilePathError, MissingColumnError
from datetime import datetime

class TestExcelProcessor(unittest.TestCase):
    def setUp(self):
        self.mock_workbook = MagicMock()
        self.mock_worksheet = MagicMock()
        self.mock_workbook.active = self.mock_worksheet
        
        # Mock file path
        self.mock_file_path = "test_file.xlsx"
        
        # Mock configuration
        self.mock_columns = [
            "Hora de Análisis",
            "Saturación (%) (Pureza)",
            "Longitud de onda (nm)"
        ]
        
        self.mock_config = {
            'columns_to_keep': self.mock_columns
        }
        
    @patch('src.excel.excel_processor.load_workbook')
    @patch('src.excel.excel_processor.validate_file_path')
    def test_init_valid_file(self, mock_validate, mock_load):
        """Test initialization with valid file path."""
        mock_validate.return_value = self.mock_file_path
        mock_load.return_value = self.mock_workbook
        
        processor = ExcelProcessor(self.mock_file_path)
        self.assertEqual(processor.file_path, self.mock_file_path)
        self.assertEqual(processor.wb, self.mock_workbook)
        self.assertEqual(processor.ws, self.mock_worksheet)

    @patch('src.excel.excel_processor.validate_file_path')
    def test_init_invalid_file(self, mock_validate):
        """Test initialization with invalid file path."""
        mock_validate.side_effect = InvalidFilePathError(self.mock_file_path)
        
        with self.assertRaises(InvalidFilePathError):
            ExcelProcessor(self.mock_file_path)

    @patch('src.excel.excel_processor.get_column_index')
    def test_validate_required_columns(self, mock_get_index):
        """Test column validation."""
        mock_processor = ExcelProcessor(self.mock_file_path)
        mock_processor.ws = self.mock_worksheet
        
        # Mock column indices
        mock_get_index.side_effect = [1, 2, 3]
        
        # Should not raise exception
        mock_processor._validate_required_columns()

    @patch('src.excel.excel_processor.get_column_index')
    def test_validate_missing_columns(self, mock_get_index):
        """Test missing column validation."""
        mock_processor = ExcelProcessor(self.mock_file_path)
        mock_processor.ws = self.mock_worksheet
        
        # Mock missing column
        mock_get_index.return_value = None
        
        with self.assertRaises(MissingColumnError):
            mock_processor._validate_required_columns()

    @patch('src.excel.excel_processor.parse_time_string')
    def test_parse_time_string(self, mock_parse):
        """Test time string parsing."""
        test_time = "14:30"
        expected_dt = datetime(2025, 1, 1, 14, 30)
        
        mock_parse.return_value = expected_dt
        
        processor = ExcelProcessor(self.mock_file_path)
        result = processor._parse_time_string(test_time)
        
        self.assertEqual(result, expected_dt)

if __name__ == '__main__':
    unittest.main()
