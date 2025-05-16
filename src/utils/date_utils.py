import pandas as pd
from datetime import datetime

def format_datetime(value):
    """Format datetime value to match the desired format."""
    if pd.isna(value):
        return None
    
    try:
        # Try parsing as datetime
        if isinstance(value, str):
            # Try different datetime formats
            formats = [
                '%d/%m/%Y %H:%M:%S',
                '%Y-%m-%d %H:%M:%S',
                '%d-%m-%Y %H:%M:%S'
            ]
            for fmt in formats:
                try:
                    dt = datetime.strptime(value, fmt)
                    return dt.strftime('%d/%m/%Y %H:%M:%S')
                except ValueError:
                    continue
        else:
            # If it's already a datetime object
            return pd.to_datetime(value).strftime('%d/%m/%Y %H:%M:%S')
    except Exception as e:
        # If parsing fails, return original value
        return value
